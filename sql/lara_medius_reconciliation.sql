with boundaries as (
   select trunc(sysdate) - 1 + 22 / 24 as end_date
     from dual
),latest as (
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
       th.haulier_code,
       pt.full_name as haulier_name,
       dep.base_country as country_code,
       th.status as to_status,
       p.trepa_to_number,
       dp_im_cost_translation.translation_code(
          c.charge_code,
          'OCEAN',
          c.transport_costs_uid
       ) as chargecode,
       c.chg_amount_fc as lineamount,
       c.self_bill_reference as selfbillingreference,
       td.bol_number as blnumber,
       td.container_number as containernumber,
       p.created_date
  from latest l
  join sce.trepa_events_payload p
on p.trepa_uid = l.trepa_uid
  join transport_header th
on th.order_no = p.trepa_to_number
  join departments dep
on dep.department_code = th.paying_agent
  join transport_costs c
on c.order_no = th.order_no
  join transport_detail_costs dc
on dc.transport_costs_uid = c.transport_costs_uid
  join transport_details td
on td.drop_no = dc.drop_no
   and c.order_no = td.order_no
  join sce.partners pt
on pt.partner_code = th.haulier_code
 order by p.trepa_to_number,
          c.charge_code,
          td.container_number,
          c.chg_amount_fc,
          c.self_bill_reference,
          td.bol_number