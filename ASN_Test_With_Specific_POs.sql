-- Updated diagnostic query to test ASN data with LIFEX as ASN number
-- This should help verify if the corrected approach finds data

-- Test: Find ASN data for specific POs from vendor performance
SELECT
    IH.LIFNR as vendor,
    IL.VGBEL as po_number,
    LTRIM(IL.MATNR, '0') as material,
    IH.VBELN as delivery_number,
    IH.LIFEX as asn_number,  -- This is the actual ASN number
    TO_DATE(IH.ERDAT,'YYYYMMDD') as creation_date,
    IH.ERNAM as created_by,
    IH.ZCARRIER as carrier,
    IH.BOLNR as bol,
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
    AND IH.VSTEL = 'IL30'
    AND IH.LFART = 'ZEL'
    AND IH.ERDAT >= '20240101'  -- Broader date range
    AND IH.LIFNR IN ('0000200000','0000210000', '200000','210000')
    AND IL.VGBEL IN ('8096470', '8097964', '8098033', '8098038', '8098039')  -- Test with specific POs from your vendor performance data
ORDER BY TO_DATE(IH.ERDAT,'YYYYMMDD') DESC
LIMIT 50;
