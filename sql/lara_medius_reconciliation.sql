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
    where created_date >= trunc(sysdate) - :lookback_days
      and created_date < trunc(sysdate)
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
outer apply
   json_table ( p.trepa_payload,'$.lines[*]'
      columns (
         chargecode varchar2 ( 150 ) path '$.chargeCode',
         lineamount number path '$.lineAmount',
         selfbillingreference varchar2 ( 150 ) path '$.selfBillingReference',
         blnumber varchar2 ( 150 ) path '$.blNumber',
         containernumber varchar2 ( 150 ) path '$.containerNumber'
      )
   )
j
 where p.rn = 1
 order by p.trepa_to_number,
          j.chargecode,
          j.containernumber,
          j.lineamount,
          j.selfbillingreference,
          j.blnumber