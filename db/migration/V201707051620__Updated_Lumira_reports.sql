-- *************************************************************************************

-- Update lumira_requests
ALTER TABLE lumira_requests ADD requestor_name varchar(255);
ALTER TABLE lumira_requests ADD owner_name varchar(255);
ALTER TABLE lumira_requests ADD current_status varchar(30);

-- Update lumira_ais
ALTER TABLE lumira_ais ADD requestor_name varchar(255);
ALTER TABLE lumira_ais ADD owner_name varchar(255);
ALTER TABLE lumira_ais ADD current_status varchar(30);

COMMIT;
