-- Comprehensive comparison to find why PO numbers don't match
-- Let's compare the actual PO number formats and ranges

-- Part 1: Sample PO numbers from Vendor Performance
SELECT 
    'Vendor_Performance_POs' as source,
    PO_NUMBER,
    VENDOR_NUMBER,
    COUNT(*) as line_count
FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
WHERE VENDOR_NUMBER IN ('200000','210000') 
AND RPT_MONTH = 'FY2025-MAY'
GROUP BY PO_NUMBER, VENDOR_NUMBER
ORDER BY PO_NUMBER
LIMIT 20

UNION ALL

-- Part 2: Sample PO numbers from ASN data (recent)
SELECT 
    'ASN_POs_Recent' as source,
    IL.VGBEL as PO_NUMBER,
    IH.LIFNR as VENDOR_NUMBER,
    COUNT(*) as line_count
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20250401'  -- Recent data
GROUP BY IL.VGBEL, IH.LIFNR
ORDER BY IL.VGBEL
LIMIT 20

UNION ALL

-- Part 3: Sample PO numbers from ASN data (broader range)
SELECT 
    'ASN_POs_Broader' as source,
    IL.VGBEL as PO_NUMBER,
    IH.LIFNR as VENDOR_NUMBER,
    COUNT(*) as line_count
FROM EDP.STD_ECC.LIKP IH 
INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
WHERE IH.MANDT = '300'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
GROUP BY IL.VGBEL, IH.LIFNR
ORDER BY IL.VGBEL
LIMIT 20;
