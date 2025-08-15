-- Test query to verify ASN data exists for vendor performance PO numbers
-- This should find matches now that we've removed the restrictive vendor/plant filters

SELECT
    IL.VGBEL as po_number,
    IH.LIFNR as asn_vendor,
    IH.LIFEX as asn_number,
    IH.ZCARRIER as carrier,
    IH.BOLNR as bol,
    TO_DATE(IH.ERDAT,'YYYYMMDD') as creation_date,
    IH.ERNAM as created_by,
    LTRIM(IL.MATNR, '0') as material,
    IL.LFIMG as quantity,
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
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'
    AND IL.VGBEL IN (
        SELECT DISTINCT PO_NUMBER 
        FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
        WHERE VENDOR_NUMBER IN ('200000','210000') 
        AND RPT_MONTH = 'FY2025-MAY'
        LIMIT 20  -- Test with first 20 PO numbers
    )
ORDER BY TO_DATE(IH.ERDAT,'YYYYMMDD') DESC
LIMIT 50;
