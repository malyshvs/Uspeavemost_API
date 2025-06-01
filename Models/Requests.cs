using System.Data.SqlClient;
using Uspevaemost_API.Controllers;

namespace Uspevaemost_API.Models
{
    public class Requests
    {
        private SqlConnection sql = new SqlConnection("Persist Security Info = False; Integrated Security=true;" +
               $"server = 10.2.9.73; Encrypt = True; TrustServerCertificate=true;");
        public Requests(string _uchps)
        {
            uchps=_uchps;
        }
        private string uchps;
        public static string uchp(string name)
        {
        SqlConnection sq = new SqlConnection("Persist Security Info = False; Integrated Security=true;" +
               $"server = 10.2.9.73; Encrypt = True; TrustServerCertificate=true;");

             string query = $@"select string_agg(''''+c.Сокращение+'''',',') 
                                    from Деканат.dbo.Пользователи a
                                    inner join Деканат.dbo.[Пользователи-Роли] b on b.КодПользователя=a.ID
                                    inner join Деканат.dbo.Факультеты c on c.Код=b.КодОбъекта
                                    inner join Деканат.dbo.Роли d on d.Код=b.кодРоли
                                    where b.КодРоли=28 and a.Логин='{name}'";
            string uchpList="";
            try
            {
                sq.Open();
                SqlCommand cmd = new SqlCommand(query, sq);
                SqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    uchpList = dr[0].ToString();
                    Console.WriteLine(dr[0].ToString());
                }
                dr.Close();

            }
            catch (SqlException er)
            {
                Console.WriteLine(er.Message);
                sq.Close();
            }
            return uchpList;
        }
        public List<string[]> getData2(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            
            string query = $@"select d.Сокращение,c.Название,c.Курс, 
max(fo.ФормаОбучения) as 'Форма обучения', 
max(uo.Уровень) as 'Уровень образования',
CONCAT(b.фамилия,' ',b.имя,' ',b.Отчество) as ФИО, 
max(b.Гражданство)as Гражданство,
max(us.Текст)as Финансирование,
max(b.льготы)as Льготы,
max(J.ekz)as Экзаменов,
max(j.sacho)as 'Зачетов с оценкой',
max(j.sach)as Зачетов,
max(j.kr)as 'Курсовых работ',
max(j.kp)as 'Курсовых проектов',
max(j.otl)as Отл,
max(J.hor) as Хор, 
max(j.tri) as Удовл,
max(j.sachet) as Зачтено,
max(j.neud) Неуд, 
max(j.nesachet) as Незачет, 
0as Абс,
0as Кач,
max(j.Стипендия)as Стипендия,
max(case when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor=0 and j.otl > 0 then'Отличник' 
when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor > 0 and j.otl >= 0 then'Хорошист' 
when j.nesachet=0 and j.neud=0 and j.tri>0 and j.hor>=0 and j.otl >= 0 then 'Успевающий с удовлетворительными оценками' 
when j.nesachet>=0 and j.neud>0 and j.tri>=0 and j.hor>=0 and j.otl >= 0 then 'Неуспевающий' 
when j.nesachet>0 and j.neud>=0 and j.tri>=0 and j.hor>=0and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость', 
max(j.сумма)as 'Сумма баллов', 
max(j.ekz + j.sach + j.sacho + j.kr + j.kp)as 'Количество дисциплин',
max(j.Среднее) as 'Средний балл', 
(max(j.tri)*3+max(j.hor)*4+max(j.otl)*5+max(j.neud)*2) as 'Сумма оценок', 
(max(j.tri)+max(j.hor)+max(j.otl)+max(j.neud)) as 'Количество оценок',
0as Ср,
max(j.neud+j.nesachet) as 'АЗ после сессии',
max(j.neud+j.nesachet - sdalsach1 - sdalp1)as 'АЗ после пересдачи 1',
max(j.neud+j.nesachet - sdalsach1 - sdalp1 - sdalsach2 - sdalp2)as 'АЗ после пересдачи 2',
max(case when j.nesachetP=0 and j.neudP=0 then 'АЗ ликвидированы'
when (j.nesachetP>=0 and j.neudP>0) or (j.nesachetP>0 and j.neudP>=0) then 'Неуспевающий'end) as 'Успеваемость после пересдач', 
max(case when b.ПродленаСессия >='{DateTime.Now}' then CONVERT(NVARCHAR,b.ПродленаСессия,23) else '' end) as 'Сессия продлена до',
max(case when per.Тип_Перемещения like'%индивидуа%' and per.ДатаПо>='{DateTime.Now}' and b.ПродленаСессия=per.ДатаПо then
Тип_Перемещения+' с '+CONVERT(NVARCHAR, ДатаС, 23)+' по '+CONVERT(NVARCHAR, ДатаПО, 23) +', приказ '+Документ else '' end) as 'Инд график'
from Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код=b.Код_Группы 
left join Деканат.dbo.Факультеты d on d.Код=c.Код_Факультета 
left join Деканат.dbo.УсловияОбучения us on us.Код=b.УслОбучения 
left join Деканат.dbo.ФормаОбучения fo on fo.Код=c.Форма_Обучения 
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи=c.Уровень
left join Деканат.dbo.Перемещения per on per.Код_Студента=b.Код
left join (select Код_Студента, Код_Группы, 
sum(Итоговый_Процент)as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz)as ekz,
sum(sach)as sach, 
sum(sacho)as sacho, 
sum(kr)as kr, 
max(Код_Ведомости)as последняяведомость,
sum(kp)as kp,
sum(sachet)as sachet,
sum(otl)as otl,
sum(hor)as hor,
sum(tri)as tri,
sum(neud)as neud, 
sum(nesachet)as nesachet,
sum(sdalsach1)as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2)as sdalp2,
sum(sdalp1)as sdalp1,
sum(sachetP)as sachetP,
sum(otlP)as otlP,
sum(horP)as horP,
sum(triP)as triP,
sum(neudP)as neudP,
sum(nesachetP)as nesachetP,
(Case when max(P.Текст) ='Коммерческое' then 'Отказ' 
when max(Оценка)=0 AND sum(otl)>0 AND sum(hor)=0 AND sum(tri)=0 then 'Отличник' 
when max(Оценка)=0 AND sum(otl)=1 AND sum(hor)>0 AND sum(tri)=0 then 'Хорошист'
when max(Оценка)=0 AND (sum(otl)>0 OR sum(hor)>0) AND sum(tri)=0 then 'Хорошист'else 'Отказ' End) as Стипендия 
from (select 
case when A.Тип_Ведомости =1 then 1 else 0 end as ekz,
case when (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))then 1 else 0 end as sach,
case when A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1 then 1 when A.Тип_Ведомости=6then 1 else 0 end as sacho, 
case when A.Тип_Ведомости=3 then 1 else 0 end as kr, case when A.Тип_Ведомости=4then 1 else 0 end as kp, 
case when (B.Итоговая_Оценка=5 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=5))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as otl, 
case when (B.Итоговая_Оценка=4 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=4))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as hor, 
case when (B.Итоговая_Оценка=3 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=3))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as tri, 
case when B.Итоговая_Оценка IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4 or A.Тип_Ведомости=12) then 1 else 0 end as neud, 
case when B.Итоговая_Оценка IN(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))and B.Пересдача1>=60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4)and B.Пересдача1>=55 then 1 else 0 end as sdalp1,
case when B.Итоговая_Оценка in(-1,1,2) and A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=0 and B.Пересдача2>=60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in(-1,1,2)and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) and B.Пересдача2>=55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка=7 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=7)then 1 else 0 end as sachet, 
case when B.Итог=5 then 1 else 0 end as otlP,
case when B.Итог=4 then 1 else 0 end as horP,
case when B.Итог=3 then 1 else 0 end as triP,
case when B.Итог IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) then 1 else 0 end as neudP,
case when B.Итог IN(-1,1,2) and A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=0 then 1 else 0 end as nesachetP,
case when B.Итог=7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN(-1,1,3,2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,Итоговый_Процент,ИтоговыйРейтинг,
Код_Студента,A.Код_Группы,Код_Ведомости, usl.Текст
from Деканат.dbo.Все_Ведомости A 
inner join Деканат.dbo.Оценки B on A.Код=B.Код_Ведомости 
left join Деканат.dbo.Все_Студенты s on B.Код_Студента=s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код=s.УслОбучения 
where (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0) and 
A.Год in({string.Join(", ", year)})and
a.Закрыта in (1)and
A.Сессия in({string.Join(", ", sem)})and A.Код_Группы=s.Код_Группы) P 
group by Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4)and
CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)like '%%'and
d.сокращение in({uchps})and
uo.Уровень in ({string.Join(", ", uo)})and
fo.ФормаОбучения in ({string.Join(", ", fo)})and
c.Курс in ({string.Join(", ", curs)})and
c.УчебныйГод in ({string.Join(", ", year)})
group by d.Сокращение, c.Название, c.Курс, CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)
order by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)";

            try
            {
                sql.Open();
                Console.WriteLine(sql.ToString());  
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[36];
                    for (int i = 0; i < 36; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        public List<string[]> getDataInv(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select d.Сокращение,c.Название,c.Курс, 
max(fo.ФормаОбучения) as 'Форма обучения', 
max(uo.Уровень) as 'Уровень образования',
CONCAT(b.фамилия,' ',b.имя,' ',b.Отчество) as ФИО, 
max(b.Гражданство)as Гражданство,
max(us.Текст)as Финансирование,
max(b.льготы)as Льготы,
max(J.ekz)as Экзаменов,
max(j.sacho)as 'Зачетов с оценкой',
max(j.sach)as Зачетов,
max(j.kr)as 'Курсовых работ',
max(j.kp)as 'Курсовых проектов',
max(j.otl)as Отл,
max(J.hor) as Хор, 
max(j.tri) as Удовл,
max(j.sachet) as Зачтено,
max(j.neud) Неуд, 
max(j.nesachet) as Незачет, 
0as Абс,
0as Кач,
max(j.Стипендия)as Стипендия,
max(case when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor=0 and j.otl > 0 then'Отличник' 
when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor > 0 and j.otl >= 0 then'Хорошист' 
when j.nesachet=0 and j.neud=0 and j.tri>0 and j.hor>=0 and j.otl >= 0 then 'Успевающий с удовлетворительными оценками' 
when j.nesachet>=0 and j.neud>0 and j.tri>=0 and j.hor>=0 and j.otl >= 0 then 'Неуспевающий' 
when j.nesachet>0 and j.neud>=0 and j.tri>=0 and j.hor>=0and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость', 
max(j.сумма)as 'Сумма баллов', 
max(j.ekz + j.sach + j.sacho + j.kr + j.kp)as 'Количество дисциплин',
max(j.Среднее) as 'Средний балл', 
(max(j.tri)*3+max(j.hor)*4+max(j.otl)*5+max(j.neud)*2) as 'Сумма оценок', 
(max(j.tri)+max(j.hor)+max(j.otl)+max(j.neud)) as 'Количество оценок',
0as Ср,
max(j.neud+j.nesachet) as 'АЗ после сессии',
max(j.neud+j.nesachet - sdalsach1 - sdalp1)as 'АЗ после пересдачи 1',
max(j.neud+j.nesachet - sdalsach1 - sdalp1 - sdalsach2 - sdalp2)as 'АЗ после пересдачи 2',
max(case when j.nesachetP=0 and j.neudP=0 then 'АЗ ликвидированы'
when (j.nesachetP>=0 and j.neudP>0) or (j.nesachetP>0 and j.neudP>=0) then 'Неуспевающий'end) as 'Успеваемость после пересдач', 
max(case when b.ПродленаСессия >='{DateTime.Now}' then CONVERT(NVARCHAR,b.ПродленаСессия,23) else '' end) as 'Сессия продлена до',
max(case when per.Тип_Перемещения like'%индивидуа%' and per.ДатаПо>='{DateTime.Now}' and b.ПродленаСессия=per.ДатаПо then
Тип_Перемещения+' с '+CONVERT(NVARCHAR, ДатаС, 23)+' по '+CONVERT(NVARCHAR, ДатаПО, 23) +', приказ '+Документ else '' end) as 'Инд график'
from Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код=b.Код_Группы 
left join Деканат.dbo.Факультеты d on d.Код=c.Код_Факультета 
left join Деканат.dbo.УсловияОбучения us on us.Код=b.УслОбучения 
left join Деканат.dbo.ФормаОбучения fo on fo.Код=c.Форма_Обучения 
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи=c.Уровень
left join Деканат.dbo.Перемещения per on per.Код_Студента=b.Код
left join (select Код_Студента, Код_Группы, 
sum(Итоговый_Процент)as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz)as ekz,
sum(sach)as sach, 
sum(sacho)as sacho, 
sum(kr)as kr, 
max(Код_Ведомости)as последняяведомость,
sum(kp)as kp,
sum(sachet)as sachet,
sum(otl)as otl,
sum(hor)as hor,
sum(tri)as tri,
sum(neud)as neud, 
sum(nesachet)as nesachet,
sum(sdalsach1)as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2)as sdalp2,
sum(sdalp1)as sdalp1,
sum(sachetP)as sachetP,
sum(otlP)as otlP,
sum(horP)as horP,
sum(triP)as triP,
sum(neudP)as neudP,
sum(nesachetP)as nesachetP,
(Case when max(P.Текст) ='Коммерческое' then 'Отказ' 
when max(Оценка)=0 AND sum(otl)>0 AND sum(hor)=0 AND sum(tri)=0 then 'Отличник' 
when max(Оценка)=0 AND sum(otl)=1 AND sum(hor)>0 AND sum(tri)=0 then 'Хорошист'
when max(Оценка)=0 AND (sum(otl)>0 OR sum(hor)>0) AND sum(tri)=0 then 'Хорошист'else 'Отказ' End) as Стипендия 
from (select 
case when A.Тип_Ведомости =1 then 1 else 0 end as ekz,
case when (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))then 1 else 0 end as sach,
case when A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1 then 1 when A.Тип_Ведомости=6then 1 else 0 end as sacho, 
case when A.Тип_Ведомости=3 then 1 else 0 end as kr, case when A.Тип_Ведомости=4then 1 else 0 end as kp, 
case when (B.Итоговая_Оценка=5 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=5))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as otl, 
case when (B.Итоговая_Оценка=4 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=4))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as hor, 
case when (B.Итоговая_Оценка=3 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=3))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as tri, 
case when B.Итоговая_Оценка IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4 or A.Тип_Ведомости=12) then 1 else 0 end as neud, 
case when B.Итоговая_Оценка IN(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10))and B.Пересдача1>=60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4)and B.Пересдача1>=55 then 1 else 0 end as sdalp1,
case when B.Итоговая_Оценка in(-1,1,2) and A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=0 and B.Пересдача2>=60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in(-1,1,2)and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) and B.Пересдача2>=55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка=7 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=7)then 1 else 0 end as sachet, 
case when B.Итог=5 then 1 else 0 end as otlP,
case when B.Итог=4 then 1 else 0 end as horP,
case when B.Итог=3 then 1 else 0 end as triP,
case when B.Итог IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) then 1 else 0 end as neudP,
case when B.Итог IN(-1,1,2) and A.Тип_Ведомости=2and A.ДиффенцированныйЗачет=0 then 1 else 0 end as nesachetP,
case when B.Итог=7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN(-1,1,3,2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,Итоговый_Процент,ИтоговыйРейтинг,
Код_Студента,A.Код_Группы,Код_Ведомости, usl.Текст
from Деканат.dbo.Все_Ведомости A 
inner join Деканат.dbo.Оценки B on A.Код=B.Код_Ведомости 
left join Деканат.dbo.Все_Студенты s on B.Код_Студента=s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код=s.УслОбучения 
where (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0) and 
A.Год in({string.Join(", ", year)})and
a.Закрыта in (1)and
A.Сессия in({string.Join(", ", sem)})and A.Код_Группы=s.Код_Группы) P 
group by Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4)and
CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)like '%%'and
d.сокращение in({uchps})and
uo.Уровень in ({string.Join(", ", uo)})and
fo.ФормаОбучения in ({string.Join(", ", fo)})and
c.Курс in ({string.Join(", ", curs)})and
c.УчебныйГод in ({string.Join(", ", year)})
and b.Льготы like '%инвалид%' and b.Льготы not like '%Родитель-инвалид%'
group by d.Сокращение, c.Название, c.Курс, CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)
order by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();

                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[36];
                    for (int i = 0; i < 36; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        public List<string[]> getInvDolgi(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = "select e.Сокращение,d.Название,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)as ФИО," +
              "b.Льготы,uo.Уровень,fo.ФормаОбучения,c.Курс,c.Код as '№ ведомости',\r\nc.Дисциплина,\r\ntv.Тип_ведомости,\r\n" +
              "c.Преподаватель,\r\nkt.Рейтинг_По_Лекциям as 'Срез 1',kt.Рейтинг_По_Практике as 'Срез 2',kt.Рейтинг_По_Лабораторным as 'Рубеж'," +
              "kt.Рейтинг_По_Другим as 'Прем',a.Надбавка as 'Экзамен',\r\na.Итоговый_Процент as 'Балл по итогу сессии',\r\n" +
              "case when a.Итоговая_Оценка = -1 then 'Незачет'\r\nwhen a.Итоговая_Оценка = 1 then 'Неявка'\r\nwhen a.Итоговая_Оценка = 2 " +
              "then 'Неуд' end as 'Оценка по итогу промежуточной аттестации',\r\ncase when isnull(a.Дата_Пересдачи1,'19330303')<>'19330303' " +
              "and isnull(a.Дата_Пересдачи2,'19330303')='19330303' then\r\nCONCAT('Пересдача 1: ',CONVERT(NVARCHAR,a.Дата_Пересдачи1,23), " +
              "', ',a.Пересдача1, ' баллов') \r\nwhen isnull(a.Дата_Пересдачи2,'19330303')<>'19330303' then\r\nCONCAT('Пересдача 1: '," +
              "CONVERT(NVARCHAR,a.Дата_Пересдачи1,23), ', ',a.Пересдача1, ' баллов; ',\r\n'Пересдача 2: ',CONVERT(NVARCHAR,a.Дата_Пересдачи2,23), " +
              "', ',a.Пересдача2, ' баллов') \r\nend,\r\na.ИтоговыйРейтинг as 'Балл после пересдачи',\r\ncase when a.Итог = -1 then 'Незачет'\r\n" +
              "when a.Итог = 1 then 'Неявка'\r\nwhen a.Итог = 2 then 'Неуд'\r\nwhen a.Итог = 7 then 'Зачет'\r\nwhen a.итог = 3 then 'Удовл'\r\n" +
              "when a.итог = 4 then 'Хор'\r\nwhen a.итог = 5 then 'Отл' end as 'Оценка после пересдачи'\r\nfrom Деканат.dbo.Оценки a\r\n" +
              "inner join Деканат.dbo.Все_Студенты b on b.Код=a.Код_Студента inner join Деканат.dbo.Все_Ведомости c on c.Код=a.Код_Ведомости\r\n" +
              "inner join Деканат.dbo.Все_Группы d on d.Код=b.Код_Группы inner join Деканат.dbo.Факультеты e on e.Код=d.Код_Факультета\r\n" +
              "left join Деканат.dbo.УсловияОбучения us on us.Код=b.УслОбучения \r\nleft join Деканат.dbo.ФормаОбучения fo on fo.Код=d.Форма_Обучения \r\n" +
              "left join Деканат.dbo.Уровень_образования uo on uo.Код_записи=d.Уровень \r\nleft join Деканат.dbo.Рейтинг_по_КТ kt on kt.Код_Оценки=a.Код\r\n" +
              "left join Деканат.dbo.Тип_Ведомости tv on tv.Код=c.Тип_Ведомости\r\nwhere  b.Льготы like '%инвалид%' and b.Льготы not like '%Родитель-инвалид%'\r\n" +
              $"and e.сокращение in({uchps})\r\n" +
              $"and uo.Уровень in ({string.Join(", ", uo)})\r\n" +
              $"and fo.ФормаОбучения in ({string.Join(", ", fo)}) \r\n" +
              $"and c.Курс in ({string.Join(", ", curs)})\r\n" +
              $"and c.Год in({string.Join(", ", year)})\r\n" +
              $"and c.Сессия in({string.Join(", ", sem)}) \r\n" +
              $"and a.Итоговая_Оценка in (-1,1,2)" +
              $" and c.Закрыта in (1) ";


            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();

                int cnt = 0;
                while (dr.Read())
                {

                    string[] row = new string[21];
                    for (int i = 0; i < 21; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    cnt++;
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }

        }
        public List<string[]> getGroups(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select Сокращение,Название, 
SUM(CASE WHEN Успеваемость='Отличник' THEN 1 ELSE 0 END)AS Отличники,
SUM(CASE WHEN Успеваемость='Хорошист' THEN 1 ELSE 0 END)AS Хорошисты,
SUM(CASE WHEN Успеваемость='Троечник' THEN 1 ELSE 0 END)AS Троечники,
SUM(CASE WHEN Успеваемость='Неуспевающий' THEN 1 ELSE 0 END)AS Неуспевающие,
sum([Сумма баллов])as 'Сумма баллов',
sum([Количество дисциплин])as 'Количество дисциплин',
sum([Сумма оценок])as 'Сумма оценок',
sum([Количество оценок])as 'Количество оценок'
from(
select d.Сокращение, c.Название, 
max(case when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor=0 and j.otl > 0 then 'Отличник' 
when j.nesachet=0 and j.neud=0 and j.tri=0 and j.hor > 0 and j.otl >= 0 then 'Хорошист' 
when j.nesachet=0 and j.neud=0 and j.tri>0 and j.hor>=0 and j.otl >= 0 then 'Троечник' 
when j.nesachet>=0 and j.neud>0 and j.tri>=0 and j.hor>=0 and j.otl >= 0 then 'Неуспевающий' 
when j.nesachet>0 and j.neud>=0 and j.tri>=0 and j.hor>=0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость', 
max(j.сумма) as 'Сумма баллов', 
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количество дисциплин',
max(j.Среднее) as 'Средний балл', 
(max(j.tri)*3+max(j.hor)*4+max(j.otl)*5+max(j.neud)*2) as 'Сумма оценок',
(max(j.tri)+max(j.hor)+max(j.otl)+max(j.neud)) as 'Количество оценок' 
from Деканат.dbo.Все_Студенты b 
left join Деканат.dbo.Все_Группы c on c.Код=b.Код_Группы 
left join Деканат.dbo.Факультеты d on d.Код=c.Код_Факультета 
left join Деканат.dbo.УсловияОбучения us on us.Код=b.УслОбучения 
left join Деканат.dbo.ФормаОбучения fo on fo.Код=c.Форма_Обучения 
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи=c.Уровень 
left join (
select Код_Студента, Код_Группы, 
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz, 
sum(sach) as sach,
sum(sacho) as sacho, 
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl, 
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP
from (
select 
case when A.Тип_Ведомости =1 then 1 else 0 end as ekz,
case when (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10)) then 1 else 0 end as sach,
case when A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1 then 1 when A.Тип_Ведомости=6 then 1 else 0 end as sacho, 
case when A.Тип_Ведомости=3 then 1 else 0 end as kr, case when A.Тип_Ведомости=4 then 1 else 0 end as kp,
case when (B.Итоговая_Оценка=5 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=5))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка=4 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=4))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка=3 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=3))
and (A.Тип_Ведомости in (1,3,4,6,12) or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4 or A.Тип_Ведомости=12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10)) then 1 else 0 end as nesachet, 
case when B.Итоговая_Оценка in(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10)) and B.Пересдача1>=60 then 1 else 0 end as sdalsach1, 
case when B.Итоговая_Оценка in(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1)
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) and B.Пересдача1>=55 then 1 else 0 end as sdalp1,
case when B.Итоговая_Оценка in(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10)) and B.Пересдача2>=60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1) 
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) and B.Пересдача2>=55 then 1 else 0 end as sdalp2, 
case when B.Итоговая_Оценка=7 or (isnull(B.Итоговая_Оценка,100)=100 and ISNULL(b.Протокол,'empty')<>'empty' and b.Итог=7) then 1 else 0 end as sachet,
case when B.Итог=5 then 1 else 0 end as otlP, 
case when B.Итог=4 then 1 else 0 end as horP, 
case when B.Итог=3 then 1 else 0 end as triP, 
case when B.Итог IN(-1,1,2) and (A.Тип_Ведомости =1 or (A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=1) 
or A.Тип_Ведомости=6 or A.Тип_Ведомости=3 or A.Тип_Ведомости=4) then 1 else 0 end as neudP, 
case when B.Итог IN(-1,1,2) and (((A.Тип_Ведомости=2 and A.ДиффенцированныйЗачет=0) or a.Тип_Ведомости=10)) then 1 else 0 end as nesachetP, 
case when B.Итог=7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN(-1,1,3,2) then 1 else 0 end) as Оценка, 
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости, 
usl.Текст 
from Деканат.dbo.Все_Ведомости A 
inner join Деканат.dbo.Оценки B on A.Код=B.Код_Ведомости 
left join Деканат.dbo.Все_Студенты s on B.Код_Студента=s.Код 
left join Деканат.dbo.УсловияОбучения usl on usl.Код=s.УслОбучения 
where (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0) and A.Год in({string.Join(", ", year)}) and a.Закрыта in (1) and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P 
group by Код_Студента,Код_Группы) J on J.Код_Студента=b.Код where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%' 
and d.сокращение in({uchps}) and uo.Уровень in ({string.Join(", ", uo)}) and fo.ФормаОбучения in ({string.Join(", ", fo)}) and c.Курс in ({string.Join(", ", curs)}) and c.УчебныйГод in ({string.Join(", ", year)}) 
group by d.Сокращение, c.Название, c.Курс, CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) ) as DataD group by Сокращение, Название order by Сокращение, Название";
            Console.WriteLine(query);
            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyUchp(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {

            string query = $@"select
Сокращение,
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок'
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by Сокращение order by Сокращение";
            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }



        public List<string[]> getbyUO(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
[Уровень образования],  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
max(uo.Уровень) as 'Уровень образования'
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by [Уровень образования] order by [Уровень образования]";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyUOCurs(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
Курс,[Уровень образования],  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
max(uo.Уровень) as 'Уровень образования',
max(c.Курс) as 'Курс'
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by Курс,[Уровень образования] order by [Уровень образования]";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[10];
                    for (int i = 0; i < 10; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyFO(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
[Форма обучения],  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
 max(fo.ФормаОбучения) as 'Форма обучения'
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by [Форма обучения] order by [Форма обучения]";


            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyCURS(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
Курс,  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
c.Курс
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by Курс order by Курс";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyQuote(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
Финансирование,  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
max(us.Текст) as 'Финансирование'
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by Финансирование order by Финансирование";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getbyCountry(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {
            string query = $@"select
Гражданство,  
SUM(CASE WHEN Успеваемость = 'Отличник' THEN 1 ELSE 0 END) AS Отличники,
SUM(CASE WHEN Успеваемость = 'Хорошист' THEN 1 ELSE 0 END) AS Хорошисты,
SUM(CASE WHEN Успеваемость = 'Троечник' THEN 1 ELSE 0 END) AS Троечники,
SUM(CASE WHEN Успеваемость = 'Неуспевающий' THEN 1 ELSE 0 END) AS Неуспевающие,
sum(Суммабаллов) as 'Сумма баллов',
sum(Количестводисциплин) as 'Количество дисциплин',
sum(Суммаоценок) as 'Сумма оценок',
sum(Количествооценок) as 'Количество оценок'
from
(select
d.Сокращение,
max(case when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor = 0 and j.otl > 0 then 'Отличник'
when j.nesachet = 0 and j.neud = 0 and j.tri = 0 and j.hor > 0 and j.otl >= 0 then 'Хорошист'
when j.nesachet = 0 and j.neud = 0 and j.tri > 0 and j.hor >= 0 and j.otl >= 0 then 'Троечник'
when j.nesachet >= 0 and j.neud > 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий'
when j.nesachet > 0 and j.neud >= 0 and j.tri >= 0 and j.hor >= 0 and j.otl >= 0 then 'Неуспевающий' end) as 'Успеваемость',
max(j.сумма) as 'Суммабаллов',
max(j.ekz + j.sach + j.sacho + j.kr + j.kp) as 'Количестводисциплин',
max(j.Среднее) as 'Средний балл',
(max(j.tri) * 3 + max(j.hor) * 4 + max(j.otl) * 5 + max(j.neud) * 2) as 'Суммаоценок',
(max(j.tri) + max(j.hor) + max(j.otl) + max(j.neud)) as 'Количествооценок',
 max(b.Гражданство) as Гражданство
from
Деканат.dbo.Все_Студенты b
left join Деканат.dbo.Все_Группы c on c.Код = b.Код_Группы
left join Деканат.dbo.Факультеты d on d.Код = c.Код_Факультета
left join Деканат.dbo.УсловияОбучения us on us.Код = b.УслОбучения
left join Деканат.dbo.ФормаОбучения fo on fo.Код = c.Форма_Обучения
left join Деканат.dbo.Уровень_образования uo on uo.Код_записи = c.Уровень
left join (select Код_Студента,Код_Группы,
sum(Итоговый_Процент) as 'Сумма',
avg(Итоговый_Процент) as 'Среднее',
sum(ekz) as ekz,
sum(sach) as sach,
sum(sacho) as sacho,
sum(kr) as kr,
max(Код_Ведомости) as последняяведомость,
sum(kp) as kp,
sum(sachet) as sachet,
sum(otl) as otl,
sum(hor) as hor,
sum(tri) as tri,
sum(neud) as neud,
sum(nesachet) as nesachet,
sum(sdalsach1) as sdalsach1,
sum(sdalsach2) as sdalsach2,
sum(sdalp2) as sdalp2,
sum(sdalp1) as sdalp1,
sum(sachetP) as sachetP,
sum(otlP) as otlP,
sum(horP) as horP,
sum(triP) as triP,
sum(neudP) as neudP,
sum(nesachetP) as nesachetP,
(Case when max(P.Текст) = 'Коммерческое' then 'Отказ'
when max(Оценка) = 0 AND sum(otl) > 0 AND sum(hor) = 0 AND sum(tri) = 0 then 'Отличник'
when max(Оценка) = 0 AND sum(otl) = 1 AND sum(hor) > 0 AND sum(tri) = 0 then 'Хорошист'
when max(Оценка) = 0 AND (sum(otl) > 0 OR sum(hor) > 0)AND sum(tri) = 0 then 'Хорошист' else 'Отказ' End) as Стипендия
from (select case when A.Тип_Ведомости = 1 then 1 else 0 end as ekz,
case when((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10) then 1 else 0 end as sach,
case
when A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1 then 1 when A.Тип_Ведомости = 6 then 1 else 0 end as sacho,
case when A.Тип_Ведомости = 3 then 1 else 0 end as kr,
case when A.Тип_Ведомости = 4 then 1 else 0 end as kp,
case when(B.Итоговая_Оценка = 5 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 5)) 
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as otl,
case when (B.Итоговая_Оценка = 4 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 4))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as hor,
case when (B.Итоговая_Оценка = 3 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty'and b.Итог = 3))
and (A.Тип_Ведомости in (1, 3, 4, 6, 12)or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)) then 1 else 0 end as tri,
case when B.Итоговая_Оценка IN (-1, 1, 2)and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1)
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 or A.Тип_Ведомости = 12) then 1 else 0 end as neud,
case when B.Итоговая_Оценка IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachet,
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10)) and B.Пересдача1 >= 60 then 1 else 0 end as sdalsach1,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача1 >= 55 then 1 else 0 end as sdalp1, 
case when B.Итоговая_Оценка in (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0) or a.Тип_Ведомости = 10))and B.Пересдача2 >= 60 then 1 else 0 end as sdalsach2,
case when B.Итоговая_Оценка in (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) 
or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4) and B.Пересдача2 >= 55 then 1 else 0 end as sdalp2,
case when B.Итоговая_Оценка = 7 or (isnull (B.Итоговая_Оценка, 100) = 100 and ISNULL (b.Протокол, 'empty') <> 'empty' and b.Итог = 7) then 1 else 0 end as sachet,
case when B.Итог = 5 then 1 else 0 end as otlP,
case when B.Итог = 4 then 1 else 0 end as horP,
case when B.Итог = 3 then 1 else 0 end as triP,
case when B.Итог IN (-1, 1, 2) and (A.Тип_Ведомости = 1 or (A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 1) or A.Тип_Ведомости = 6 or A.Тип_Ведомости = 3 or A.Тип_Ведомости = 4 ) then 1 else 0 end as neudP,
case when B.Итог IN (-1, 1, 2) and (((A.Тип_Ведомости = 2 and A.ДиффенцированныйЗачет = 0)or a.Тип_Ведомости = 10)) then 1 else 0 end as nesachetP,
case when B.Итог = 7 then 1 else 0 end as sachetP,
(Case when B.Итоговая_Оценка IN (-1, 1, 3, 2) then 1 else 0 end) as Оценка,
Итоговая_Оценка,
Итоговый_Процент,
ИтоговыйРейтинг,
Код_Студента,
A.Код_Группы,
Код_Ведомости,
usl.Текст
from
Деканат.dbo.Все_Ведомости A
inner join Деканат.dbo.Оценки B on A.Код = B.Код_Ведомости
left join Деканат.dbo.Все_Студенты s on B.Код_Студента = s.Код
left join Деканат.dbo.УсловияОбучения usl on usl.Код = s.УслОбучения
where  (B.Оценка_По_Рейтингу!=6 or isnull(B.Скрыта,0)=0)
and A.Год in({string.Join(", ", year)})
and a.Закрыта in (1)
and A.Сессия in({string.Join(", ", sem)}) and A.Код_Группы=s.Код_Группы) P
group by  Код_Студента,Код_Группы) J on J.Код_Студента=b.Код
where b.Статус in (1,4) and CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество) like '%%'       
and d.сокращение in({uchps})
and uo.Уровень in ({string.Join(", ", uo)})
and fo.ФормаОбучения in ({string.Join(", ", fo)})
and c.Курс in ({string.Join(", ", curs)})
and c.УчебныйГод in ({string.Join(", ", year)})  
group by d.Сокращение, c.Название, c.Курс,CONCAT(b.фамилия, ' ',b.имя,' ',b.Отчество)) as DataD  
group by Гражданство order by Гражданство";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[9];
                    for (int i = 0; i < 9; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }
        public List<string[]> getDisc(List<string> year, List<string> sem, List<string> uo, List<string> fo, List<string> curs)
        {

            string query = "select b.код,concat (c.Фамилия,' ', c.Имя,' ', c.Отчество) as ФИО, e.Сокращение, d.Название, d.Курс, c.Льготы, " +
                "b.Дисциплина,concat(e.сокращение,' ',b.Дисциплина),concat(e.сокращение,' ',b.Преподаватель), b.Преподаватель," +
                "concat(left(n.ОКСО,2),'.00.00')as УГСН,n.ОКСО, n.Название_Спец as 'Наименование НПС', p.Название_Спец as 'Направление (профиль)', a.Итоговый_Процент," +
                " case when a.Итоговая_Оценка = 7 then 'Зачет' when a.Итоговая_Оценка = 1 then 'Неявка' when a.Итоговая_Оценка = -1 then 'Незачет' else concat(a.Итоговая_Оценка,'')" +
                " end as Оценка, a.Оценка_По_Рейтингу,t.Тип_ведомости, b.ДиффенцированныйЗачет,b.Закрыта, uo.Уровень,fo.ФормаОбучения,st.Описание, " +
                "concat(n.оксо,' ',n.Название_Спец) as 'Код+НПС', concat(d.Курс,' ',uo.Уровень) as 'Курс+Уровень' " +
                "from Деканат.dbo.Оценки a " +
                "left join Деканат.dbo.Все_Ведомости b on b.код=a.Код_Ведомости " +
                "left join Деканат.dbo.Все_Студенты c on c.код=a.Код_Студента " +
                "left join Деканат.dbo.все_группы d on d.код=c.Код_Группы " +
                "left join Деканат.dbo.Факультеты e on e.код=d.Код_Факультета " +
                "left join Деканат.dbo.Специальности n on n.Код=d.Код_Специальности " +
                "left join Деканат.dbo.Специальности p on p.код=d.Код_Профиль " +
                "inner join Деканат.dbo.Тип_Ведомости t on t.Код=b.Тип_Ведомости " +
                "inner join Деканат.dbo.Уровень_образования uo on uo.Код_записи=d.Уровень " +
                "inner join Деканат.dbo.ФормаОбучения fo on fo.Код=d.Форма_Обучения " +
                $"inner join Деканат.dbo.Статус_Студента st on st.Код=c.Статус " +
                $"where b.Год in ({string.Join(", ", year)}) and d.Курс in ({string.Join(", ", curs)}) and e.Сокращение in ({uchps})" +
                $"and uo.Уровень in ({string.Join(", ", uo)}) and b.сессия in ({string.Join(", ", sem)}) and fo.ФормаОбучения in ({string.Join(", ", fo)})" +
                $"and b.Закрыта in (1)";

            try
            {
                sql.Open();
                List<string[]> list = new List<string[]>();
                SqlCommand cmd = new SqlCommand(query, sql);
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string[] row = new string[25];
                    for (int i = 0; i < 25; i++)
                    {
                        row[i] = dr[i].ToString();
                    }
                    list.Add(row);
                }
                dr.Close();
                sql.Close();
                return list;

            }
            catch (SqlException e)
            {
                Console.WriteLine(e.Message);
                sql.Close();
                return null;
            }
        }

    }
}
