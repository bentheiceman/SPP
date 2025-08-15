-- Advanced ASN Discovery Query
-- This query tries multiple smart strategies to link ASN data with vendor performance data

WITH vendor_analysis AS (
    SELECT 
        'Vendor Performance Vendors' as source,
        VENDOR_NUMBER,
        VENDOR_NAME,
        COUNT(*) as record_count
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE VENDOR_NUMBER IN ('200000','210000') AND RPT_MONTH = 'FY2025-MAY'
    GROUP BY VENDOR_NUMBER, VENDOR_NAME
    
    UNION ALL
    
    SELECT 
        'ASN System Vendors' as source,
        IH.LIFNR as VENDOR_NUMBER,
        V.NAME1 as VENDOR_NAME,
        COUNT(*) as record_count
    FROM EDP.STD_ECC.LIKP IH 
    INNER JOIN EDP.STD_ECC.LFA1 V ON IH.MANDT = V.MANDT AND IH.LIFNR = V.LIFNR
    WHERE IH.MANDT = '300' AND IH.LFART = 'ZEL' AND IH.ERDAT >= '20240101'
    GROUP BY IH.LIFNR, V.NAME1
),

material_analysis AS (
    SELECT 
        'Vendor Performance Materials' as source,
        USN as material_number,
        COUNT(DISTINCT PO_NUMBER) as po_count,
        COUNT(*) as line_count
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE VENDOR_NUMBER IN ('200000','210000') AND RPT_MONTH = 'FY2025-MAY'
        AND USN IS NOT NULL
    GROUP BY USN
    
    UNION ALL
    
    SELECT 
        'ASN System Materials' as source,
        LTRIM(IL.MATNR, '0') as material_number,
        COUNT(DISTINCT IL.VGBEL) as po_count,
        COUNT(*) as line_count
    FROM EDP.STD_ECC.LIKP IH 
    INNER JOIN EDP.STD_ECC.LIPS IL ON IH.VBELN = IL.VBELN AND IH.MANDT = IL.MANDT
    WHERE IH.MANDT = '300' AND IH.LFART = 'ZEL' AND IH.ERDAT >= '20240101'
        AND IL.MATNR IS NOT NULL
    GROUP BY LTRIM(IL.MATNR, '0')
),

material_matches AS (
    SELECT 
        vp.material_number as vp_material,
        vp.po_count as vp_po_count,
        asn.material_number as asn_material,
        asn.po_count as asn_po_count,
        'Exact Material Match' as match_type
    FROM material_analysis vp
    INNER JOIN material_analysis asn ON vp.material_number = asn.material_number
    WHERE vp.source = 'Vendor Performance Materials' 
        AND asn.source = 'ASN System Materials'
),

recent_asn_activity AS (
    SELECT 
        IH.LIFNR as vendor_number,
        V.NAME1 as vendor_name,
        COUNT(DISTINCT IH.LIFEX) as unique_asn_count,
        COUNT(*) as total_records,
        MIN(TO_DATE(IH.ERDAT,'YYYYMMDD')) as earliest_date,
        MAX(TO_DATE(IH.ERDAT,'YYYYMMDD')) as latest_asn_date,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.LIFEX), '; ') as sample_asn_numbers,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.ZCARRIER), '; ') as sample_carriers,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.BOLNR), '; ') as sample_bol_numbers
    FROM EDP.STD_ECC.LIKP IH 
    INNER JOIN EDP.STD_ECC.LFA1 V ON IH.MANDT = V.MANDT AND IH.LIFNR = V.LIFNR
    WHERE IH.MANDT = '300' AND IH.LFART = 'ZEL' 
        AND IH.ERDAT >= '20240501'
        AND IH.LIFEX IS NOT NULL
    GROUP BY IH.LIFNR, V.NAME1
    HAVING COUNT(DISTINCT IH.LIFEX) >= 5
),

vendor_name_matches AS (
    SELECT 
        vp.VENDOR_NUMBER as vp_vendor,
        vp.VENDOR_NAME as vp_name,
        asn.vendor_number as asn_vendor,
        asn.vendor_name as asn_name,
        asn.unique_asn_count,
        asn.sample_asn_numbers,
        asn.sample_carriers,
        asn.sample_bol_numbers,
        CASE 
            WHEN UPPER(TRIM(vp.VENDOR_NAME)) = UPPER(TRIM(asn.vendor_name)) THEN 'Exact Name Match'
            WHEN UPPER(vp.VENDOR_NAME) LIKE '%' || UPPER(TRIM(asn.vendor_name)) || '%' THEN 'VP Contains ASN Name'
            WHEN UPPER(TRIM(asn.vendor_name)) LIKE '%' || UPPER(vp.VENDOR_NAME) || '%' THEN 'ASN Contains VP Name'
            ELSE 'No Match'
        END as match_type
    FROM vendor_analysis vp
    CROSS JOIN recent_asn_activity asn
    WHERE vp.source = 'Vendor Performance Vendors'
),

expanded_asn_search AS (
    SELECT 
        IH.LIFNR as vendor_number,
        V.NAME1 as vendor_name,
        V.STRAS as vendor_address,
        V.ORT01 as vendor_city,
        COUNT(DISTINCT IH.LIFEX) as asn_count,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.LIFEX), '; ') as asn_numbers,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.ZCARRIER), '; ') as carriers,
        ARRAY_TO_STRING(ARRAY_AGG(DISTINCT IH.BOLNR), '; ') as bol_numbers,
        MAX(TO_DATE(IH.ERDAT,'YYYYMMDD')) as latest_asn_date
    FROM EDP.STD_ECC.LIKP IH 
    INNER JOIN EDP.STD_ECC.LFA1 V ON IH.MANDT = V.MANDT AND IH.LIFNR = V.LIFNR
    WHERE IH.MANDT = '300' AND IH.LFART = 'ZEL' 
        AND IH.ERDAT >= '20240101'
        AND IH.LIFEX IS NOT NULL
        AND (
            UPPER(V.NAME1) LIKE '%SUPPLIER%' OR
            UPPER(V.NAME1) LIKE '%VENDOR%' OR
            UPPER(V.NAME1) LIKE '%CORP%' OR
            UPPER(V.NAME1) LIKE '%INC%' OR
            UPPER(V.NAME1) LIKE '%LLC%' OR
            UPPER(V.NAME1) LIKE '%LTD%'
        )
    GROUP BY IH.LIFNR, V.NAME1, V.STRAS, V.ORT01
    HAVING COUNT(DISTINCT IH.LIFEX) >= 3
)

SELECT 
    section_order,
    analysis_section,
    detail_1,
    detail_2,
    detail_3,
    detail_4,
    detail_5
FROM (
    SELECT 
        '1_VENDOR_ANALYSIS' as section_order,
        'VENDOR NAME ANALYSIS' as analysis_section,
        match_type as detail_1,
        CONCAT(vp_vendor, ': ', LEFT(vp_name, 30)) as detail_2,
        CONCAT(asn_vendor, ': ', LEFT(asn_name, 30)) as detail_3,
        unique_asn_count::VARCHAR as detail_4,
        LEFT(sample_asn_numbers, 50) as detail_5
    FROM vendor_name_matches 
    WHERE match_type != 'No Match'
    
    UNION ALL
    
    SELECT 
        '2_MATERIAL_ANALYSIS' as section_order,
        'MATERIAL MATCHES' as analysis_section,
        vp_material as detail_1,
        vp_po_count::VARCHAR as detail_2,
        asn_po_count::VARCHAR as detail_3,
        match_type as detail_4,
        '' as detail_5
    FROM material_matches
    
    UNION ALL
    
    SELECT 
        '3_ASN_VENDORS' as section_order,
        'TOP ASN VENDORS (Recent Activity)' as analysis_section,
        CONCAT(vendor_number, ': ', LEFT(vendor_name, 25)) as detail_1,
        unique_asn_count::VARCHAR as detail_2,
        latest_asn_date::VARCHAR as detail_3,
        LEFT(sample_asn_numbers, 40) as detail_4,
        LEFT(sample_carriers, 25) as detail_5
    FROM recent_asn_activity
    
    UNION ALL
    
    SELECT 
        '4_EXPANDED_SEARCH' as section_order,
        'EXPANDED ASN SEARCH' as analysis_section,
        CONCAT(vendor_number, ': ', LEFT(vendor_name, 25)) as detail_1,
        asn_count::VARCHAR as detail_2,
        LEFT(asn_numbers, 40) as detail_3,
        LEFT(carriers, 25) as detail_4,
        latest_asn_date::VARCHAR as detail_5
    FROM expanded_asn_search
)
ORDER BY section_order, detail_2;
