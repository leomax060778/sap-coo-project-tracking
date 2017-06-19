

-- *************************************************************************************

-- Update lumira_requests
ALTER TABLE lumira_requests ADD original_due datetime;
ALTER TABLE lumira_requests ADD deleted bit default 0;
ALTER TABLE lumira_requests ADD deleted_date datetime;
ALTER TABLE lumira_requests RENAME in_progess TO in_progress;

-- Update lumira_ais
ALTER TABLE lumira_ais ADD approved	datetime;
ALTER TABLE lumira_ais ADD deleted bit default 0;
ALTER TABLE lumira_ais ADD deleted_date datetime;
ALTER TABLE lumira_ais ADD original_owner varchar(128);
ALTER TABLE lumira_ais ADD delivery_rejected_reason varchar(255);
ALTER TABLE lumira_ais ALTER extension datetime;

COMMIT;