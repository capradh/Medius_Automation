with boundaries as (
   select trunc(sysdate) + 22 / 24
          - case
               when sysdate < trunc(sysdate) + 22 / 24 then 1
               else 0
            end as end_date
     from dual
),
latest as (
   select trepa_uid
     from (
      select tep.trepa_uid,
             row_number()
             over(partition by tep.trepa_to_number
                  order by tep.created_date desc,
                           tep.trepa_uid desc
             ) as rn
        from sce.trepa_events_payload tep
       where tep.created_date >= (
                 select end_date - :lookback_days
                   from boundaries
             )
         and tep.created_date < (
                 select end_date
                   from boundaries
             )
   )
    where rn = 1
)
select p.trepa_uid,
       j.haulier_code,
       pt.full_name as haulier_name,
       d.base_country as country_code,
       th.status as to_status,
       p.trepa_to_number,
       j.chargecode,
       j.lineamount,
       j.selfbillingreference,
       j.blnumber,
       j.containernumber,
       p.created_date
  from latest l
  join sce.trepa_events_payload p
on p.trepa_uid = l.trepa_uid
  join transport_header th
on th.order_no = p.trepa_to_number
  join departments d
on d.department_code = th.paying_agent
cross apply
   json_table ( p.trepa_payload,'$'
      columns (
         haulier_code varchar2 ( 150 ) path '$.supplierCode',
         nested path '$.lines[*]'
            columns (
               chargecode varchar2 ( 150 ) path '$.chargeCode',
               lineamount number path '$.lineAmount',
               selfbillingreference varchar2 ( 150 ) path '$.selfBillingReference',
               blnumber varchar2 ( 150 ) path '$.blNumber',
               containernumber varchar2 ( 150 ) path '$.containerNumber'
            )
      )
   )
j
  join sce.partners pt
on pt.partner_code = j.haulier_code
 order by p.trepa_to_number,
          j.chargecode,
          j.containernumber,
          j.lineamount,
          j.selfbillingreference,
          j.blnumber