-- Enhanced Vendor Performance Query 
-- ASN integration removed due to different data systems

WITH primary_metric AS (
    SELECT
        VENDOR_NUMBER,
        VENDOR_NAME,
        CONCAT(VENDOR_NUMBER,' - ',vendor_name) AS VENDOR,
        PO_NUMBER,
        USN,
        ITEM_DESCRIPTION,
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        CONCAT (USN, ' - ', ITEM_DESCRIPTION) as SKU,
        TO_DATE(DATE_ORIG_ORDERED) AS DATE_ORIG_ORDERED,
        TO_DATE(DATE_ORIG_PROMISED) AS DATE_ORIG_PROMISED,
        TO_DATE(DATE_FIRST_RECEIVED) AS DATE_FIRST_RECEIVED,
        WAREHOUSE_NUM,
        WAREHOUSE_NAME,
        METRIC,
        RPT_MONTH,
        FSCL_YR_PRD,
        METRIC_NUMERATOR,
        METRIC_DENOMINATOR,
        NETWORK
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE 
        VENDOR_NUMBER IN ('200000','210000')
        AND RPT_MONTH = 'FY2025-MAY'
        AND METRIC IN ('First_Receipt_FR_B1D', 'First_Receipt_FR_B28D')
),

-- Fill Rate 28D calculation
fill_rate_28d AS (
    SELECT
        VENDOR_NUMBER,
        PO_NUMBER,
        USN,
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        SUM(CASE WHEN METRIC = 'First_Receipt_FR_B28D' THEN METRIC_NUMERATOR ELSE 0 END) AS units_received_28d,
        SUM(CASE WHEN METRIC = 'First_Receipt_FR_B28D' THEN METRIC_DENOMINATOR ELSE 0 END) AS units_ordered_28d,
        CASE 
            WHEN SUM(CASE WHEN METRIC = 'First_Receipt_FR_B28D' THEN METRIC_DENOMINATOR ELSE 0 END) > 0 
            THEN ROUND((SUM(CASE WHEN METRIC = 'First_Receipt_FR_B28D' THEN METRIC_NUMERATOR ELSE 0 END) * 100.0) / 
                       SUM(CASE WHEN METRIC = 'First_Receipt_FR_B28D' THEN METRIC_DENOMINATOR ELSE 0 END), 2)
            ELSE 0 
        END AS fill_rate_28d_pct
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE 
        VENDOR_NUMBER IN ('200000','210000')
        AND RPT_MONTH = 'FY2025-MAY'
        AND METRIC = 'First_Receipt_FR_B28D'
    GROUP BY VENDOR_NUMBER, PO_NUMBER, USN
),

-- Shipments in Full (Fill Rate 1D)
shipments_in_full AS (
    SELECT
        VENDOR_NUMBER,
        PO_NUMBER,
        USN,
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_NUMERATOR ELSE 0 END) AS units_received_1d,
        SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_DENOMINATOR ELSE 0 END) AS units_ordered_1d,
        CASE 
            WHEN SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_DENOMINATOR ELSE 0 END) > 0 
            THEN ROUND((SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_NUMERATOR ELSE 0 END) * 100.0) / 
                       SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_DENOMINATOR ELSE 0 END), 2)
            ELSE 0 
        END AS shipments_in_full_pct,
        CASE 
            WHEN SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_NUMERATOR ELSE 0 END) = 
                 SUM(CASE WHEN METRIC = 'First_Receipt_FR_B1D' THEN METRIC_DENOMINATOR ELSE 0 END)
            THEN 'Complete'
            ELSE 'Partial'
        END AS shipment_status
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE 
        VENDOR_NUMBER IN ('200000','210000')
        AND RPT_MONTH = 'FY2025-MAY'
        AND METRIC = 'First_Receipt_FR_B1D'
    GROUP BY VENDOR_NUMBER, PO_NUMBER, USN
),

-- Units On-Time Complete calculation
units_ontime_complete AS (
    SELECT
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        PO_NUMBER,
        USN,
        VENDOR_NUMBER,
        DATE_ORIG_PROMISED,
        DATE_FIRST_RECEIVED,
        METRIC_NUMERATOR as units_received,
        METRIC_DENOMINATOR as units_ordered,
        CASE 
            WHEN DATE_FIRST_RECEIVED <= DATE_ORIG_PROMISED 
                 AND METRIC_NUMERATOR = METRIC_DENOMINATOR 
            THEN METRIC_NUMERATOR
            ELSE 0
        END AS units_ontime_complete,
        CASE 
            WHEN DATE_FIRST_RECEIVED <= DATE_ORIG_PROMISED 
                 AND METRIC_NUMERATOR = METRIC_DENOMINATOR 
            THEN 'On-Time Complete'
            WHEN DATE_FIRST_RECEIVED <= DATE_ORIG_PROMISED 
                 AND METRIC_NUMERATOR < METRIC_DENOMINATOR
            THEN 'On-Time Partial'
            WHEN DATE_FIRST_RECEIVED > DATE_ORIG_PROMISED 
                 AND METRIC_NUMERATOR = METRIC_DENOMINATOR
            THEN 'Late Complete'
            ELSE 'Late Partial'
        END AS delivery_status
    FROM DM_SUPPLYCHAIN.VENDOR_PERFORMANCE.COMBINED_IPR_IB_VENDOR_PERFORMANCE
    WHERE 
        VENDOR_NUMBER IN ('200000','210000')
        AND RPT_MONTH = 'FY2025-MAY'
        AND METRIC IN ('First_Receipt_FR_B1D', 'First_Receipt_FR_B28D')
),

hds_receipts AS (
    SELECT 
        CONCAT(ebeln, ':', LTRIM(MATNR, '0')) AS Metric_Concatenate,
        MAX(TRY_TO_DATE(TO_CHAR(budat), 'yyyymmdd')) AS receipt_date
    FROM edp.std_ecc.ekbe
    WHERE bwart IN ('101', '102')
    GROUP BY ebeln, MATNR
),

hdp_receipts AS (
    SELECT
        CONCAT(PO_NUMBER, ':', USN) AS Metric_Concatenate,
        MAX(TO_DATE(DATE_RECEIVED)) AS receipt_date
    FROM DM_SUPPLYCHAIN.PRO_INVENTORY_ANALYTICS.REPORT_PURCHASE_ORDER_VISIBILITY_SHIPMENTS
    GROUP BY PO_NUMBER, USN
),

combined_receipts AS (
    SELECT
        Metric_Concatenate,
        MAX(receipt_date) AS receipt_date
    FROM (
        SELECT * FROM hds_receipts
        UNION ALL
        SELECT * FROM hdp_receipts
    ) all_receipts
    GROUP BY Metric_Concatenate
)

-- Final SELECT with all metrics
SELECT
    pm.RPT_MONTH as Report_Month,
    pm.NETWORK,
    pm.VENDOR,
    pm.WAREHOUSE_NUM,
    pm.WAREHOUSE_NAME,
    pm.PO_NUMBER,
    pm.SKU,
    
    -- ASN Information (Not available - different systems)
    'ASN Data Not Available - Different System' AS "ASN_Status",
    'No ASN Data' AS "ASN_Numbers", 
    'No Carrier Data' AS "Carriers",
    'No BOL Data' AS "BOL_Numbers",
    
    pm.DATE_ORIG_ORDERED as Date_Ordered,
    pm.DATE_ORIG_PROMISED as Date_Promised,
    pm.DATE_FIRST_RECEIVED,
    cr.receipt_date as Date_Last_Received,
    pm.METRIC,
    zeroifnull(pm.METRIC_NUMERATOR) as Metric_Units_Received,
    pm.METRIC_DENOMINATOR as Metric_Units_Ordered,
    
    -- Original compliance calculation
    case 
        when Metric_Units_Received < Metric_Units_Ordered
        then 'Non-Compliant'
        else 'Compliant'
    end as "Result",
    
    -- Fill-Rate (28D)
    COALESCE(fr28.fill_rate_28d_pct, 0) AS "Fill_Rate_28D_Percent",
    COALESCE(fr28.units_received_28d, 0) AS "Units_Received_28D",
    COALESCE(fr28.units_ordered_28d, 0) AS "Units_Ordered_28D",
    
    -- Shipments in Full (Fill Rate 1D)
    COALESCE(sif.shipments_in_full_pct, 0) AS "Shipments_In_Full_Percent",
    COALESCE(sif.shipment_status, 'Unknown') AS "Shipment_Status",
    COALESCE(sif.units_received_1d, 0) AS "Units_Received_1D",
    COALESCE(sif.units_ordered_1d, 0) AS "Units_Ordered_1D",
    
    -- Units On-Time Complete
    COALESCE(uotc.units_ontime_complete, 0) AS "Units_OnTime_Complete",
    COALESCE(uotc.delivery_status, 'Unknown') AS "Delivery_Status",
    CASE 
        WHEN uotc.units_ordered > 0 
        THEN ROUND((uotc.units_ontime_complete * 100.0) / uotc.units_ordered, 2)
        ELSE 0 
    END AS "OnTime_Complete_Percent",
    
    -- ASN Utilization Metrics (Not available)
    0 AS "ASN_Utilization_Percent",
    0 AS "Shipments_With_ASN",
    0 AS "Total_Shipments_ASN",
    0 AS "Unique_ASN_Count"
    
FROM primary_metric pm
LEFT JOIN combined_receipts cr
    ON pm.Metric_Concatenate = cr.Metric_Concatenate
LEFT JOIN fill_rate_28d fr28
    ON pm.Metric_Concatenate = fr28.Metric_Concatenate
LEFT JOIN shipments_in_full sif
    ON pm.Metric_Concatenate = sif.Metric_Concatenate
LEFT JOIN units_ontime_complete uotc
    ON pm.Metric_Concatenate = uotc.Metric_Concatenate

ORDER BY pm.VENDOR, pm.PO_NUMBER, pm.USN;
