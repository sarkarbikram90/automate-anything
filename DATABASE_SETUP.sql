-- ========================================================================
-- Excel to MySQL Automation - Database Setup Script
-- ========================================================================
-- Run this script in MySQL Workbench or any MySQL client
-- to prepare your database for the automation
-- ========================================================================

-- ========================================================================
-- STEP 1: CREATE DATABASE
-- ========================================================================

-- Create database (change 'excel_import_db' to your preferred name)
CREATE DATABASE IF NOT EXISTS excel_import_db;

-- Select the database
USE excel_import_db;

-- ========================================================================
-- STEP 2: CREATE TABLE (BASIC EXAMPLE)
-- ========================================================================

-- Option A: Simple table with common columns
-- Modify columns to match your actual Excel structure
CREATE TABLE IF NOT EXISTS excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    
    -- Your actual columns (MODIFY THESE based on your Excel file)
    name VARCHAR(255),
    email VARCHAR(255) UNIQUE,
    phone VARCHAR(20),
    amount DECIMAL(10, 2),
    department VARCHAR(100),
    
    -- Metadata columns (automatically populated)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    
    -- Indexing for better performance
    INDEX idx_email (email),
    INDEX idx_created_at (created_at)
);

-- ========================================================================
-- STEP 3: CREATE ADDITIONAL INDEXES (OPTIONAL)
-- ========================================================================

-- Add indexes if querying by specific columns
CREATE INDEX idx_name ON excel_data(name);
CREATE INDEX idx_department ON excel_data(department);

-- ========================================================================
-- STEP 4: CREATE AUDIT/BACKUP TABLE (OPTIONAL)
-- ========================================================================

-- Keep a historical record of imports
CREATE TABLE IF NOT EXISTS excel_data_audit (
    audit_id INT AUTO_INCREMENT PRIMARY KEY,
    action VARCHAR(50),  -- INSERT, UPDATE, DELETE
    record_id INT,
    old_data JSON,
    new_data JSON,
    modified_by VARCHAR(255),
    modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- ========================================================================
-- STEP 5: VERIFY SETUP
-- ========================================================================

-- Show databases
SHOW DATABASES;

-- Show tables in current database
SHOW TABLES;

-- Show table structure
DESCRIBE excel_data;

-- ========================================================================
-- STEP 6: CREATE DATABASE USER (OPTIONAL BUT RECOMMENDED)
-- ========================================================================

-- Create a dedicated user for the automation
-- This is more secure than using root account
-- Change 'automation_user' and 'secure_password' to your values

-- Create user
CREATE USER IF NOT EXISTS 'automation_user'@'localhost' IDENTIFIED BY 'secure_password';

-- Grant permissions on specific database
GRANT SELECT, INSERT, UPDATE, DELETE ON excel_import_db.* TO 'automation_user'@'localhost';

-- For remote connections (if needed):
-- CREATE USER 'automation_user'@'%' IDENTIFIED BY 'secure_password';
-- GRANT SELECT, INSERT, UPDATE, DELETE ON excel_import_db.* TO 'automation_user'@'%';

-- Refresh privileges
FLUSH PRIVILEGES;

-- Verify user creation
SELECT User, Host FROM mysql.user WHERE User = 'automation_user';

-- ========================================================================
-- STEP 7: EXAMPLE DATA STRUCTURE
-- ========================================================================

-- Example 1: Basic employee data
/*
CREATE TABLE IF NOT EXISTS excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    employee_id VARCHAR(20) UNIQUE,
    first_name VARCHAR(100),
    last_name VARCHAR(100),
    email VARCHAR(255) UNIQUE,
    phone VARCHAR(20),
    department VARCHAR(100),
    position VARCHAR(100),
    salary DECIMAL(10, 2),
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
*/

-- Example 2: Sales data
/*
CREATE TABLE IF NOT EXISTS excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    sales_order_id VARCHAR(20) UNIQUE,
    customer_name VARCHAR(255),
    product_name VARCHAR(255),
    quantity INT,
    unit_price DECIMAL(10, 2),
    total_amount DECIMAL(12, 2),
    order_date DATE,
    delivery_date DATE,
    status VARCHAR(50),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
*/

-- Example 3: Contact information
/*
CREATE TABLE IF NOT EXISTS excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    contact_name VARCHAR(255),
    company VARCHAR(255),
    email VARCHAR(255) UNIQUE,
    phone VARCHAR(20),
    mobile VARCHAR(20),
    address TEXT,
    city VARCHAR(100),
    state VARCHAR(100),
    zip_code VARCHAR(10),
    country VARCHAR(100),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
*/

-- ========================================================================
-- STEP 8: TEST DATA (OPTIONAL - for testing)
-- ========================================================================

-- Insert test data to verify table structure
INSERT INTO excel_data (name, email, phone, amount, department) 
VALUES 
    ('John Doe', 'john@example.com', '+1-555-0001', 5000.00, 'Sales'),
    ('Jane Smith', 'jane@example.com', '+1-555-0002', 4500.00, 'Marketing'),
    ('Bob Johnson', 'bob@example.com', '+1-555-0003', 6000.00, 'Engineering');

-- Verify data insertion
SELECT * FROM excel_data;

-- ========================================================================
-- CLEANUP/RESET (USE WITH CAUTION!)
-- ========================================================================

-- To clear all data but keep table structure:
-- TRUNCATE TABLE excel_data;

-- To drop the entire table:
-- DROP TABLE IF EXISTS excel_data;

-- To drop the entire database:
-- DROP DATABASE IF EXISTS excel_import_db;

-- ========================================================================
-- CONFIGURATION FOR AUTOMATION SCRIPT
-- ========================================================================

-- Use these values in your .env file:
-- MYSQL_HOST=localhost
-- MYSQL_USER=automation_user          (or root if preferred)
-- MYSQL_PASSWORD=secure_password      (use actual password)
-- MYSQL_DATABASE=excel_import_db      (database name created above)
-- MYSQL_PORT=3306
-- TABLE_NAME=excel_data               (table name created above)

-- ========================================================================
-- MONITORING/MAINTENANCE QUERIES
-- ========================================================================

-- Check table size
SELECT 
    TABLE_NAME,
    ROUND(((data_length + index_length) / 1024 / 1024), 2) AS 'Size (MB)'
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'excel_import_db';

-- Count total records
SELECT COUNT(*) as total_records FROM excel_data;

-- Find duplicate emails
SELECT email, COUNT(*) FROM excel_data 
GROUP BY email HAVING COUNT(*) > 1;

-- Find records from last 24 hours
SELECT * FROM excel_data 
WHERE created_at >= DATE_SUB(NOW(), INTERVAL 1 DAY);

-- ========================================================================
-- BACKUP/EXPORT (from command line)
-- ========================================================================

-- Backup entire database:
-- mysqldump -u root -p excel_import_db > backup_2024-01-15.sql

-- Backup specific table:
-- mysqldump -u root -p excel_import_db excel_data > backup_excel_data.sql

-- Export to CSV:
-- SELECT * FROM excel_data INTO OUTFILE '/tmp/export.csv' 
-- FIELDS TERMINATED BY ',' ENCLOSED BY '"';

-- ========================================================================
-- END OF SETUP SCRIPT
-- ========================================================================
