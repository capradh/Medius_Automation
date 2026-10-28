WITH po_lines AS (
    SELECT
        p.po_identifier_sq,
        p.call_id,
        p.supplier_code,
        p.last_update_date,
        j.main_currency,
        j.charge_code,
        j.line_amount,
        j.line_ordinal
    FROM crac.vcr_po_identifiers p
    LEFT OUTER JOIN JSON_TABLE(
        CASE
            WHEN p.last_publish_payload IS JSON THEN p.last_publish_payload
            ELSE TO_CLOB('{"lines":[]}')
        END,
        '$.lines[*]'
        COLUMNS (
            line_ordinal FOR ORDINALITY,
            line_amount   NUMBER        PATH '$.lineAmount' NULL ON ERROR,
            main_currency VARCHAR2(150) PATH '$.mainCurrency' NULL ON ERROR,
            charge_code   VARCHAR2(150) PATH '$.chargeCode' NULL ON ERROR
        )
    ) j ON 1 = 1
    WHERE p.last_update_date >= TRUNC(SYSDATE) - :lookback_days
      AND p.last_update_date < TRUNC(SYSDATE)
)
SELECT
    po_identifier_sq,
    call_id,
    supplier_code,
    main_currency AS "mainCurrency",
    charge_code AS "chargeCode",
    last_update_date,
    po_identifier_sq || '-' || charge_code AS "uniqueVal",
    line_amount AS "totalLineAmount",
    line_ordinal
FROM po_lines
ORDER BY
    po_identifier_sq DESC,
    charge_code NULLS LAST,
    last_update_date DESC,
    call_id,
    line_ordinal
