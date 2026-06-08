SELECT current_database(), current_schema(), inet_server_addr(), inet_server_port();
SELECT COUNT(*) AS policy_rows FROM policydata_detail_raw;
SELECT COUNT(*) AS current_members FROM app_policydata_current_members;
SELECT COUNT(*) AS relation_summary_rows FROM app_policy_relation_summary;
SELECT * FROM app_policy_relation_summary ORDER BY franchise_name, relation LIMIT 100;
