-- Alternative approach: Check if there might be a vendor mapping table
-- or if we need to approach this differently

-- Check if there are any ASN records for vendors that might map to 200000/210000
-- Look for pattern similarities or potential mappings

-- Part 1: Look for vendors with similar patterns to 200000/210000
SELECT DISTINCT
    'Potential_Vendor_Match' as test_type,
    IH.LIFNR as asn_vendor,
    COUNT(DISTINCT IL.VGBEL) as po_count,
    MIN(IL.VGBEL) as min_po,
    MAX(IL.VGBEL) as max_po,
    COUNT(*) as total_lines
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
    AND (
        IH.LIFNR LIKE '%200000%' OR 
        IH.LIFNR LIKE '%210000%' OR
        IH.LIFNR LIKE '0000200000' OR 
        IH.LIFNR LIKE '0000210000'
    )
GROUP BY IH.LIFNR
ORDER BY po_count DESC;

-- Part 2: Check date ranges to see if there's a timing mismatch
SELECT 
    'ASN_Date_Distribution' as test_type,
    EXTRACT(YEAR FROM TO_DATE(IH.ERDAT,'YYYYMMDD')) as year,
    EXTRACT(MONTH FROM TO_DATE(IH.ERDAT,'YYYYMMDD')) as month,
    COUNT(*) as record_count,
    COUNT(DISTINCT IL.VGBEL) as unique_pos
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
GROUP BY EXTRACT(YEAR FROM TO_DATE(IH.ERDAT,'YYYYMMDD')), 
         EXTRACT(MONTH FROM TO_DATE(IH.ERDAT,'YYYYMMDD'))
ORDER BY year DESC, month DESC
LIMIT 12;

-- Part 3: Check if PO numbers might have prefixes or different formats
SELECT 
    'PO_Number_Analysis' as test_type,
    LENGTH(IL.VGBEL) as po_length,
    SUBSTR(IL.VGBEL, 1, 2) as po_prefix,
    COUNT(DISTINCT IL.VGBEL) as unique_pos,
    MIN(IL.VGBEL) as min_po_example,
    MAX(IL.VGBEL) as max_po_example
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
GROUP BY LENGTH(IL.VGBEL), SUBSTR(IL.VGBEL, 1, 2)
ORDER BY unique_pos DESC
LIMIT 20;
