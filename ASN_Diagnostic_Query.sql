-- Diagnostic query to test ASN data availability
-- Run this first to verify ASN data exists for your POs

-- Test 1: Check if ASN data exists for any of the POs from vendor performance
SELECT 
    'PO_List_from_Vendor_Performance' as test_type,
    COUNT(*) as record_count,
    COUNT(DISTINCT PO_NUMBER) as unique_pos
FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
WHERE VENDOR_NUMBER IN ('200000','210000') 
AND RPT_MONTH = 'FY2025-MAY';

-- Test 2: Check what PO numbers exist in vendor performance
SELECT DISTINCT 
    PO_NUMBER,
    VENDOR_NUMBER,
    COUNT(*) as line_count
FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
WHERE VENDOR_NUMBER IN ('200000','210000') 
AND RPT_MONTH = 'FY2025-MAY'
GROUP BY PO_NUMBER, VENDOR_NUMBER
ORDER BY PO_NUMBER
LIMIT 10;

-- Test 3: Check if ASN data exists for these vendors (broader date range)
SELECT
    'ASN_Data_Check' as test_type,
    COUNT(*) as total_records,
    COUNT(DISTINCT IH.VBELN) as unique_deliveries,
    COUNT(DISTINCT IL.VGBEL) as unique_pos,
    MIN(TO_DATE(IH.ERDAT,'YYYYMMDD')) as earliest_date,
    MAX(TO_DATE(IH.ERDAT,'YYYYMMDD')) as latest_date
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.VSTEL = 'IL30'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20250101'  -- Broader date range
    AND IH.LIFNR IN ('0000200000','0000210000', '200000','210000');

-- Test 4: Sample ASN data to see format
SELECT
    IH.LIFNR as vendor,
    IL.VGBEL as po_number,
    LTRIM(IL.MATNR, '0') as material,
    IH.VBELN as delivery_number,
    TO_DATE(IH.ERDAT,'YYYYMMDD') as creation_date,
    IH.ERNAM as created_by,
    IH.ZCARRIER as carrier,
    IH.BOLNR as bol,
    CASE 
        WHEN IH.VBELN LIKE '06%' AND IH.ERNAM = 'BPAREMOTE' THEN 'ASN'
        WHEN IH.VBELN LIKE '06%' AND IH.ERNAM = 'SCEBATCH' THEN 'ASN'
        WHEN IH.VBELN LIKE '06%' AND IH.ERNAM = 'P2P_IDOC' THEN 'ASN'
        WHEN IH.VBELN LIKE '06%' AND IH.ERNAM = 'P2PBATCH' THEN 'ASN'	
        WHEN IH.VBELN LIKE '10%' THEN 'EGR'
        ELSE 'Manually Created'
    END AS inbound_type
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.VSTEL = 'IL30'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20250101'
    AND IH.LIFNR IN ('0000200000','0000210000', '200000','210000')
ORDER BY TO_DATE(IH.ERDAT,'YYYYMMDD') DESC
LIMIT 20;
