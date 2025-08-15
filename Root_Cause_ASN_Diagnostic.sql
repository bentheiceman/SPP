-- Comprehensive diagnostic to find why ASN data isn't matching
-- Let's check each component separately

-- Step 1: Check what vendor numbers actually exist in ASN tables
SELECT DISTINCT 
    'ASN_Vendor_Check' as test_type,
    IH.LIFNR as vendor_number,
    COUNT(*) as record_count
FROM EDP.STD_ECC.LIKP IH 
WHERE IH.MANDT = '300'
    AND IH.VSTEL = 'IL30'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20230101'  -- Very broad date range
GROUP BY IH.LIFNR
ORDER BY record_count DESC
LIMIT 20;

-- Step 2: Check what plants/warehouses exist in ASN data
SELECT DISTINCT 
    'ASN_Plant_Check' as test_type,
    IH.VSTEL as plant,
    COUNT(*) as record_count
FROM EDP.STD_ECC.LIKP IH 
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20230101'
GROUP BY IH.VSTEL
ORDER BY record_count DESC
LIMIT 10;

-- Step 3: Check date ranges in ASN data
SELECT 
    'ASN_Date_Range' as test_type,
    MIN(TO_DATE(IH.ERDAT,'YYYYMMDD')) as earliest_date,
    MAX(TO_DATE(IH.ERDAT,'YYYYMMDD')) as latest_date,
    COUNT(*) as total_records
FROM EDP.STD_ECC.LIKP IH 
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL';

-- Step 4: Check if any PO numbers from vendor performance exist in ASN data (any vendor, any plant)
SELECT
    'PO_Match_Check' as test_type,
    COUNT(DISTINCT vp.PO_NUMBER) as vp_pos,
    COUNT(DISTINCT IL.VGBEL) as asn_pos,
    COUNT(DISTINCT CASE WHEN IL.VGBEL = vp.PO_NUMBER THEN vp.PO_NUMBER END) as matching_pos
FROM (
    SELECT DISTINCT PO_NUMBER 
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE VENDOR_NUMBER IN ('200000','210000') 
    AND RPT_MONTH = 'FY2025-MAY'
) vp
FULL OUTER JOIN (
    SELECT DISTINCT IL.VGBEL
    FROM EDP.STD_ECC.LIKP IH 
    INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
    WHERE IH.MANDT = '300'
        AND IH.LFART = 'ZEL'
        AND IH.ERDAT >= '20230101'
) asn_data ON vp.PO_NUMBER = asn_data.VGBEL;

-- Step 5: Sample of actual PO numbers in ASN data
SELECT 
    'Sample_ASN_POs' as test_type,
    IL.VGBEL as po_number,
    IH.LIFNR as vendor,
    COUNT(*) as line_count
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
GROUP BY IL.VGBEL, IH.LIFNR
ORDER BY COUNT(*) DESC
LIMIT 20;
