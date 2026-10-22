with latest_payload as (
   select trepa_uid,
          trepa_to_number,
          trepa_payload,
          created_date,
          row_number()
          over(partition by trepa_to_number
               order by created_date desc,
                        trepa_uid desc
          ) as rn
     from sce.trepa_events_payload
    where created_month >= trunc(
         sysdate - :lookback_days,
         'MM'
      )
      and created_date >= sysdate - :lookback_days
      and trepa_payload is json
)
select p.trepa_uid,
       p.trepa_to_number,
       j.chargecode,
       j.lineamount,
       j.selfbillingreference,
       j.blnumber,
       j.containernumber,
       p.created_date
  from latest_payload p
  left outer join
   json_table ( p.trepa_payload,'$'
      columns (
         nested path '$.lines[*]'
            columns (
               chargeCode varchar2 ( 150 ) path '$.chargeCode' null on error,
               lineAmount number path '$.lineAmount' null on error,
               selfBillingReference varchar2 ( 150 ) path '$.selfBillingReference' null on error,
               blNumber varchar2 ( 150 ) path '$.blNumber' null on error,
               containerNumber varchar2 ( 150 ) path '$.containerNumber' null on error
            )
      )
   )
j
on 1 = 1
 where p.rn = 1
 order by p.trepa_to_number,
          j.chargecode,
          j.containernumber,
          j.lineamount,
          j.selfbillingreference,
          j.blnumber