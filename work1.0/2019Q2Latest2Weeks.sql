use CSA_2019Q2

select * from P4P_temp
select * from NP_temp
select * from Infeeds_temp
select * from P4P_temp1

select distinct 客户 from P4P_20190514
select distinct 广告主 from P4P_20190514

--drop table P4P_temp
--drop table NP_temp
--drop table Infeeds_temp
--drop table P4P_temp1
--drop table P4P_temp2

select 用户名, 广告主, 客户, AM, channel, [2019Apr]+[2019May]+[2019Jun] as [Q2QTD_P4P_Spending],
[20190501]+[20190502]+[20190503]+[20190504]+[20190505]+[20190506]+[20190507] as [0501_0507_P4P_Spending],
[20190508]+[20190509]+[20190510]+[20190511]+[20190512]+[20190513]+[20190514] as [0508_0514_P4P_Spending]
into P4P_temp
from P4P_20190514
where 端口 not like '%wrong%'
order by 用户名, 广告主, 客户, AM

select 用户名, 广告主, 客户, AM, channel, [2019Apr]+[2019May]+[2019Jun] as [Q2QTD_NP_Spending],
[20190501]+[20190502]+[20190503]+[20190504]+[20190505]+[20190506]+[20190507] as [0501_0507_NP_Spending],
[20190508]+[20190509]+[20190510]+[20190511]+[20190512]+[20190513]+[20190514] as [0508_0514_NP_Spending]
into NP_temp
from NP_20190514
where 端口 not like '%wrong%'
order by 用户名, 广告主, 客户, AM

select 用户名, 广告主, 客户, AM, channel, [2019Apr]+[2019May]+[2019Jun] as [Q2QTD_Infeeds_Spending],
[20190501]+[20190502]+[20190503]+[20190504]+[20190505]+[20190506]+[20190507] as [0501_0507_Infeeds_Spending],
[20190508]+[20190509]+[20190510]+[20190511]+[20190512]+[20190513]+[20190514] as [0508_0514_Infeeds_Spending]
into Infeeds_temp
from Infeeds_20190514
where 端口 not like '%wrong%'
order by 用户名, 广告主, 客户, AM

-- Fergus version 链接查询P4P & NP & Infeeds
select a.*, b.Q2QTD_NP_Spending, b.[0501_0507_NP_Spending], b.[0508_0514_NP_Spending], 
c.Q2QTD_Infeeds_Spending, c.[0501_0507_Infeeds_Spending], c.[0508_0514_Infeeds_Spending]
into P4P_temp1
from P4P_temp a
left join NP_temp b
on a.用户名 = b.用户名
left join Infeeds_temp c
on b.用户名 = c.用户名

-- Brian version 合并P4P & NP & Infeeds   --ignore
alter table P4P_temp add Q2QTD_NP_Spending float
update P4P_temp
set Q2QTD_NP_Spending = b.Q2QTD_NP_Spending
from P4P_temp a inner join NP_temp b on a.用户名=b.用户名

alter table P4P_temp add [0501_0507_NP_Spending] float
update P4P_temp
set [0501_0507_NP_Spending] = b.[0501_0507_NP_Spending]
from P4P_temp a inner join NP_temp b on a.用户名=b.用户名

alter table P4P_temp add [0508_0514_NP_Spending] float
update P4P_temp
set [0508_0514_NP_Spending] = b.[0508_0514_NP_Spending]
from P4P_temp a inner join NP_temp b on a.用户名=b.用户名

alter table P4P_temp add [Q2QTD_Infeeds_Spending] float
update P4P_temp
set [Q2QTD_Infeeds_Spending] = b.[Q2QTD_Infeeds_Spending]
from P4P_temp a inner join Infeeds_temp b on a.用户名=b.用户名

alter table P4P_temp add [0501_0507_Infeeds_Spending] float
update P4P_temp
set [0501_0507_Infeeds_Spending] = b.[0501_0507_Infeeds_Spending]
from P4P_temp a inner join Infeeds_temp b on a.用户名=b.用户名

alter table P4P_temp add [0508_0514_Infeeds_Spending] float
update P4P_temp
set [0508_0514_Infeeds_Spending] = b.[0508_0514_Infeeds_Spending]
from P4P_temp a inner join Infeeds_temp b on a.用户名=b.用户名

--Consolidating Master for Output
select 广告主, 客户, AM, channel, sum([Q2QTD_P4P_Spending])+sum([Q2QTD_NP_Spending])+sum([Q2QTD_Infeeds_Spending]) as [Q2QTD_Total_Spending], 
sum([0501_0507_P4P_Spending])+sum([0501_0507_NP_Spending])+sum([0501_0507_Infeeds_Spending]) as [0501_0507_Total_Spending],
sum([0508_0514_P4P_Spending])+sum([0508_0514_NP_Spending])+sum([0508_0514_Infeeds_Spending]) as [0508_0514_Total_Spending],
sum([0501_0507_P4P_Spending]) as [0501_0507_P4P_Spending],
sum([0508_0514_P4P_Spending]) as [0508_0514_P4P_Spending],
sum([0501_0507_NP_Spending])+sum([0501_0507_Infeeds_Spending]) as [0501_0507_NewProduct_Spending],
sum([0508_0514_NP_Spending])+sum([0508_0514_Infeeds_Spending]) as [0508_0514_NewProduct_Spending],
sum([0501_0507_P4P_Spending])+sum([0501_0507_NP_Spending])+sum([0501_0507_Infeeds_Spending])+sum([0508_0514_P4P_Spending])+sum([0508_0514_NP_Spending])+sum([0508_0514_Infeeds_Spending]) as [0501_0514_Total_Spending]
into P4P_temp2
from P4P_temp1
group by 广告主, 客户, AM, channel
order by [0501_0514_Total_Spending] desc

--TOP50广告主消费
select top 80 rank() over (order by sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) desc) as 序号, 广告主, 客户, AM, 
sum(Q2QTD_Total_Spending) as [Q2QTD_Total_Spending], sum([0501_0507_Total_Spending]) as [0501_0507_Total_Spending] , 
sum([0508_0514_Total_Spending]) as [0508_0514_Total_Spending], sum([0501_0507_P4P_Spending]) as [0501_0507_P4P_Spending], 
sum([0508_0514_P4P_Spending]) as [0508_0514_P4P_Spending] , sum([0501_0507_NewProduct_Spending]) as [0501_0507_NewProduct_Spending], 
sum([0508_0514_NewProduct_Spending]) as [0508_0514_NewProduct_Spending],
rank() over(order by sum([0501_0507_Total_Spending]) desc) 上周排名,
sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]) as [Diff_Total_Spending],
case when sum([0501_0507_Total_Spending])=0 then 0 else
(sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]))/sum([0501_0507_Total_Spending]) end as [Diff%_Total_Spending],
sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) as [0501_0514_Total_Spending]
from P4P_temp2
group by 广告主, 客户, AM
order by [0501_0514_Total_Spending] desc

--近两周代理商消费
select rank() over (order by sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) desc) as Rank, 客户, channel, 
sum([Q2QTD_Total_Spending]) as [Q2QTD_Total_Spending], sum([0501_0507_Total_Spending]) as [0501_0507_Total_Spending] , 
sum([0508_0514_Total_Spending]) as [0508_0514_Total_Spending], sum([0501_0507_P4P_Spending]) as [0501_0507_P4P_Spending], 
sum([0508_0514_P4P_Spending]) as [0508_0514_P4P_Spending] , sum([0501_0507_NewProduct_Spending]) as [0501_0507_NewProduct_Spending], 
sum([0508_0514_NewProduct_Spending]) as [0508_0514_NewProduct_Spending],
sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]) as [Diff_Total_Spending],
case when sum([0501_0507_Total_Spending])=0 then 0 else 
(sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]))/sum([0501_0507_Total_Spending]) end as [Diff%_Total_Spending],
sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) as [0501_0514_Total_Spending]
from P4P_temp2
where [Q2QTD_Total_Spending]>0
group by 客户, channel
order by [0501_0514_Total_Spending] desc

--近两周广告主消费
select rank() over (order by sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) desc) as Rank, 广告主, 客户, AM, 
sum([Q2QTD_Total_Spending]) as [Q2QTD_Total_Spending], sum([0501_0507_Total_Spending]) as [0501_0507_Total_Spending] , 
sum([0508_0514_Total_Spending]) as [0508_0514_Total_Spending], sum([0501_0507_P4P_Spending]) as [0501_0507_P4P_Spending], 
sum([0508_0514_P4P_Spending]) as [0508_0514_P4P_Spending] , sum([0501_0507_NewProduct_Spending]) as [0501_0507_NewProduct_Spending], 
sum([0508_0514_NewProduct_Spending]) as [0508_0514_NewProduct_Spending],
sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]) as [Diff_Total_Spending],
case when sum([0501_0507_Total_Spending])=0 then 0 else 
(sum([0508_0514_Total_Spending])-sum([0501_0507_Total_Spending]))/sum([0501_0507_Total_Spending]) end as [Diff%_Total_Spending],
sum([0501_0507_Total_Spending])+sum([0508_0514_Total_Spending]) as [0501_0514_Total_Spending]
from P4P_temp2
where [Q2QTD_Total_Spending]>0
group by 广告主, 客户, AM
order by [0501_0514_Total_Spending] desc