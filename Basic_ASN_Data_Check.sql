-- Simplified ASN search - remove restrictive filters to see if any data exists

-- Very basic ASN data check
SELECT 
    COUNT(*) as total_likp_records,
    COUNT(DISTINCT VBELN) as unique_deliveries,
    MIN(TO_DATE(ERDAT,'YYYYMMDD')) as earliest_date,
    MAX(TO_DATE(ERDAT,'YYYYMMDD')) as latest_date
FROM EDP.STD_ECC.LIKP 
WHERE MANDT = '300';

-- Check what document types exist
SELECT DISTINCT 
    LFART as document_type,
    COUNT(*) as record_count
FROM EDP.STD_ECC.LIKP 
WHERE MANDT = '300'
GROUP BY LFART
ORDER BY record_count DESC;

-- Check what plants exist
SELECT DISTINCT 
    VSTEL as plant,
    COUNT(*) as record_count
FROM EDP.STD_ECC.LIKP 
WHERE MANDT = '300'
    AND LFART IN ('ZEL', 'EL', 'LF', 'NL')  -- Common delivery types
GROUP BY VSTEL
ORDER BY record_count DESC
LIMIT 10;

-- Sample ASN records without restrictive filters
SELECT 
    LIFNR as vendor,
    VSTEL as plant,
    VBELN as delivery_number,
    LIFEX as asn_number,
    TO_DATE(ERDAT,'YYYYMMDD') as creation_date,
    ERNAM as created_by,
    LFART as doc_type,
    ZCARRIER as carrier,
    BOLNR as bol
FROM EDP.STD_ECC.LIKP 
WHERE MANDT = '300'
    AND ERDAT >= '20240101'
    AND LIFEX IS NOT NULL  -- Only records with ASN numbers
ORDER BY TO_DATE(ERDAT,'YYYYMMDD') DESC
LIMIT 20;
