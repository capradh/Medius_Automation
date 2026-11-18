with boundaries as (
   select trunc(sysdate) + 22 / 24 -
          case
             when sysdate < trunc(sysdate) + 22 / 24 then
                1
             else
                0
          end
   as end_date
     from dual
),po_lines as (
   select p.po_identifier_sq,
          p.call_id,
          p.supplier_code,
          p.country_code,
          p.last_update_date,
          j.main_currency,
          j.charge_code,
          j.line_amount,
          j.line_ordinal
     from crac.vcr_po_identifiers p
    cross join boundaries b
     left outer join
      json_table (
         case
            when p.last_publish_payload is json then
               p.last_publish_payload
            else
               to_clob('{
    "lines": []
}')
         end,'$.lines[*]'
         columns (
            line_ordinal for ordinality,
            line_amount number path '$.lineAmount' null on error,
            main_currency varchar2 ( 150 ) path '$.mainCurrency' null on error,
            charge_code varchar2 ( 150 ) path '$.chargeCode' null on error
         )
      )
   j
   on 1 = 1
    where p.last_update_date >= b.end_date - :lookback_days
      and p.last_update_date < b.end_date
)
select po_identifier_sq,
       call_id,
       supplier_code,
       country_code,
       main_currency as "mainCurrency",
       charge_code as "chargeCode",
       last_update_date,
       po_identifier_sq
       || '-'
       || charge_code as "uniqueVal",
       line_amount as "totalLineAmount",
       line_ordinal
  from po_lines
 order by po_identifier_sq desc,
          charge_code nulls last,
          last_update_date desc,
          call_id,
          line_ordinal