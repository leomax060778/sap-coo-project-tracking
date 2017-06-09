

-- *************************************************************************************
-- Update schema version
--INSERT INTO SCHEMA_VERSION(VERSION, DESCRIPTION, SCRIPT) VALUES('V1.0.0-01', 'Add sent notification field to system', 'V201703161700__Add_sent_notification.sql');

ALTER TABLE system ADD sent_notification DATE;

-- Set initial date
UPDATE system SET sent_notification = '2017-03-01' WHERE id = 1;


COMMIT;