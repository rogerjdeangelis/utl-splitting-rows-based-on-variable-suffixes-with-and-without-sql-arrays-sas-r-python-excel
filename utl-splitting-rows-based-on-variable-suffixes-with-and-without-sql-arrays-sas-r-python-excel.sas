%let pgm=utl-splitting-rows-based-on-variable-suffixes-with-and-without-sql-arrays-sas-r-python-excel;

%stop_submission;

Splitting rows based on variable suffixes with and without sql arrays sas r python excel

    CONTENTS

       1 sas sql no arrays
       2 sas sql arrays
       3 r sql recursive
       4 python sql
         for excel see https://tinyurl.com/4e6yaap8
        SOAP BOX ON
         Took me over an hour to convert the series datatype from first sql call to a string that
         I could use  the in the second sql call. Maybe I should not use  triple quotes in the first sql call.
         I fially got it to work using, which was not needed in r.
         txt=cde['cde'].astype(str).str.cat(sep=' ')
        SOAPBOX OFF


github
https://tinyurl.com/25sx38x9
https://github.com/rogerjdeangelis/utl-splitting-rows-based-on-variable-suffixes-with-and-without-sql-arrays-sas-r-python-excel

communities.sas
https://tinyurl.com/35zjhert
https://communities.sas.com/t5/SAS-Programming/Sincerely-asking-A-question-related-to-proc-transpose/m-p/815308

/**************************************************************************************************************************/
/*                                |                                               |                                       */
/*            INPUT               |           PROCESS                             |         OUTPUT                        */
/*            =====               |           ========                            |         ======                        */
/*                                |                                               |                                       */
/*                                |  Create two rows for each input row           |  PA PB       SUFFIX A B  C            */
/*  PA PB      A1 A2 B1 B2 C1 C2  |  One row    with PA PB and A1 B1 C1           |                                       */
/*                                |  Second row with PA PB and A2 B2 C2           |  A  fail        6   1 8  1            */
/*  A  success  1  2  3  4  5  6  |                                               |  A  fail        1   5 7  9            */
/*  A  fail     5  6  7  8  9  1  |                                               |  A  success     2   1 4  6            */
/*  B  success  9  1  2  3  4  5  |  1 SAS NO ARRAYS                              |  A  success     1   1 3  5            */
/*  B  fail     4  5  6  7  8  9  |  ===============                              |  B  fail        5   1 7  9            */
/*  C  success  8  9  1  2  3  4  |                                               |  B  fail        1   4 6  8            */
/*  C  fail     3  4  5  6  7  8  |  proc sql;                                    |  B  success     1   1 3  5            */
/*                                |   create                                      |  B  success     1   9 2  4            */
/*                                |     table want as                             |  C  fail        4   1 6  8            */
/*  options validvarname=upcase;  |   select                                      |  C  fail        1   3 5  7            */
/*  libname sd1 "d:/sd1";         |     pa                                        |  C  success     9   1 2  4            */
/*  data sd1.have;                |    ,pb                                        |  C  success     1   8 1  3            */
/*    input Pa $ pb $             |    ,1 as suffix                               |                                       */
/*     a1 a2 b1 b2 c1 c2 ;        |    ,a1 as a                                   |                                       */
/*  cards4;                       |    ,b1 as b                                   |                                       */
/*  A  success 1 2 3 4 5 6        |    ,c1 as c                                   |                                       */
/*  A  fail    5 6 7 8 9 1        |   from                                        |                                       */
/*  B  success 9 1 2 3 4 5        |     sd1.have                                  |                                       */
/*  B  fail    4 5 6 7 8 9        |   union                                       |                                       */
/*  C  success 8 9 1 2 3 4        |     all                                       |                                       */
/*  C  fail    3 4 5 6 7 8        |   select                                      |                                       */
/*  ;;;;                          |     pa                                        |                                       */
/*  run;quit;                     |    ,pb                                        |                                       */
/*                                |    ,1 as suffix                               |                                       */
/*                                |    ,a2 as a                                   |                                       */
/*                                |    ,b2 as b                                   |                                       */
/*                                |    ,c2 as c                                   |                                       */
/*                                |   from                                        |                                       */
/*                                |     sd1.have                                  |                                       */
/*                                |   order                                       |                                       */
/*                                |     by pa, pb, a                              |                                       */
/*                                |  ;quit;                                       |                                       */
/*                                |                                               |                                       */
/*                                |---------------------------------------------------------------------------------------*/
/*                                |                                               |                                       */
/*                                |  2 sas sql arrays                             |                                       */
/*                                |  ================                             |  PA PB       SUFFIX A B  C            */
/*                                |                                               |                                       *
/*                                |   %array(_sfx,values=1 2);                    |  A  fail        6   1 8  1            */
/*                                |                                               |  A  fail        1   5 7  9            */
/*                                |   proc sql;                                   |  A  success     2   1 4  6            */
/*                                |     create                                    |  A  success     1   1 3  5            */
/*                                |       table want as                           |  B  fail        5   1 7  9            */
/*                                |       %do_over(_sfx,phrase=%str(              |  B  fail        1   4 6  8            */
/*                                |         select                                |  B  success     1   1 3  5            */
/*                                |           pa                                  |  B  success     1   9 2  4            */
/*                                |          ,pb                                  |  C  fail        4   1 6  8            */
/*                                |          ,? as suffix                         |  C  fail        1   3 5  7            */
/*                                |          ,a? as a                             |  C  success     9   1 2  4            */
/*                                |          ,b? as b                             |  C  success     1   8 1  3            */
/*                                |          ,c? as c                             |                                       */
/*                                |         from                                  |                                       */
/*                                |           have), between=union all)           |                                       */
/*                                |         order                                 |                                       */
/*                                |           by pa, pb, suffix                   |                                       */
/*                                |   ;quit;                                      |                                       */
/*                                |                                               |                                       */
/*                                |    Highlight the code and save                |                                       */
/*                                |    to clipboard and then submit               |                                       */
/*                                |    macro debugx and the generated             |                                       */
/*                                |    code will be in the log                    |                                       */
/*                                |                                               |                                       */
/*                                |---------------------------------------------------------------------------------------*/
/*                                |                                               |                                       */
/*                                |    3 R SQL RECURSION                          |   R                                   */
/*                                |    =================                          |                                       */
/*                                |                                               |   PA     PB suffix a b c              */
/*                                |    Note the fist call to sql generates        |   A success      1 1 3 5              */
/*                                |    the sql code, while the second sql         |   A    fail      1 5 7 9              */
/*                                |    executes the generated code.               |   B success      1 9 2 4              */
/*                                |    By minimizing contact with base python     |   B    fail      1 4 6 8              */
/*                                |    we have code that works in mutiple sql     |   C success      1 8 1 3              */
/*                                |    dialects.                                  |   C    fail      1 3 5 7              */
/*                                |                                               |   A success      2 2 4 6              */
/*                                |    proc datasets lib=sd1 nolist nodetails;    |   A    fail      2 6 8 1              */
/*                                |     delete want;                              |   B success      2 1 3 5              */
/*                                |    run;quit;                                  |   B    fail      2 5 7 9              */
/*                                |                                               |   C success      2 9 2 4              */
/*                                |    %utl_rbeginx;                              |   C    fail      2 4 6 8              */
/*                                |    parmcards4;                                |                                       */
/*                                |    library(haven)                             |   SAS                                 */
/*                                |    library(sqldf)                             |   PA PB       SUFFIX A B  C           */
/*                                |    source("c:/oto/fn_tosas9x.R")              |                                       */
/*                                |    options(sqldf.dll = "d:/dll/sqlean.dll")   |   A  fail        6   1 8  1           */
/*                                |    have<-read_sas("d:/sd1/have.sas7bdat")     |   A  fail        1   5 7  9           */
/*                                |    print(have)                                |   A  success     2   1 4  6           */
/*                                |    cde<-sqldf('                               |   A  success     1   1 3  5           */
/*                                |     with                                      |   B  fail        5   1 7  9           */
/*                                |          recursive cnt(x) as (                |   B  fail        1   4 6  8           */
/*                                |       select                                  |   B  success     1   1 3  5           */
/*                                |          1                                    |   B  success     1   9 2  4           */
/*                                |       union                                   |   C  fail        4   1 6  8           */
/*                                |          all                                  |   C  fail        1   3 5  7           */
/*                                |       select                                  |   C  success     9   1 2  4           */
/*                                |          x+1                                  |   C  success     1   8 1  3           */
/*                                |       from                                    |                                       */
/*                                |          cnt                                  |                                       */
/*                                |       where                                   |                                       */
/*                                |          x < 2                                |                                       */
/*                                |       ),                                      |                                       */
/*                                |     qs as                                     |                                       */
/*                                |         (select                               |                                       */
/*                                |           "select                             |                                       */
/*                                |              pa                               |                                       */
/*                                |             ,pb                               |                                       */
/*                                |             ,?  as suffix                     |                                       */
/*                                |             ,a?  as a                         |                                       */
/*                                |             ,b?  as b                         |                                       */
/*                                |             ,c?  as c                         |                                       */
/*                                |          from                                 |                                       */
/*                                |             have" as str)                     |                                       */
/*                                |     select                                    |                                       */
/*                                |       group_concat(replace(                   |                                       */
/*                                |          r.str                                |                                       */
/*                                |         ,"?"                                  |                                       */
/*                                |         ,cast(l.x as text))                   |                                       */
/*                                |         ," union all ") as cde                |                                       */
/*                                |     from                                      |                                       */
/*                                |         cnt as l left join qs as r            |                                       */
/*                                |     on 1=1                                    |                                       */
/*                                |    ')                                         |                                       */
/*                                |    cde                                        |                                       */
/*                                |    want<-sqldf(cde$cde)                       |                                       */
/*                                |    want                                       |                                       */
/*                                |    fn_tosas9x(                                |                                       */
/*                                |          inp    = want                        |                                       */
/*                                |         ,outlib ="d:/sd1/"                    |                                       */
/*                                |         ,outdsn ="want"                       |                                       */
/*                                |         )                                     |                                       */
/*                                |    ;;;;                                       |                                       */
/*                                |    %utl_rendx;                                |                                       */
/*                                |                                               |                                       */
/*                                |    proc print data=sd1.want;                  |                                       */
/*                                |    run;quit;                                  |                                       */
/*                                |                                               |                                       */
/*                                |---------------------------------------------------------------------------------------*/
/*                                |                                               |                                       */
/*                                | 3 PYTHON SQL RECURSION                        | Python                                */
/*                                | =======================                       |    PA       PB  suffix  a    b    c   */
/*                                |                                               | 0   A  success    1   1.0  3.0  5.0   */
/*                                | proc datasets lib=sd1 nolist nodetails;       | 1   A     fail    1   5.0  7.0  9.0   */
/*                                |  delete pywant;                               | 2   B  success    1   9.0  2.0  4.0   */
/*                                | run;quit;                                     | 3   B     fail    1   4.0  6.0  8.0   */
/*                                |                                               | 4   C  success    1   8.0  1.0  3.0   */
/*                                | %utl_pybeginx;                                | 5   C     fail    1   3.0  5.0  7.0   */
/*                                | parmcards4;                                   | 6   A  success    2   2.0  4.0  6.0   */
/*                                | import re                                     | 7   A     fail    2   6.0  8.0  1.0   */
/*                                | exec(open('c:/oto/fn_pythonx.py').read());    | 8   B  success    2   1.0  3.0  5.0   */
/*                                | have,meta = ps.read_sas7bdat( \               | 9   B     fail    2   5.0  7.0  9.0   */
/*                                |   'd:/sd1/have.sas7bdat');                    | 10  C  success    2   9.0  2.0  4.0   */
/*                                | cde=pdsql('''                                 | 11  C     fail    2   4.0  6.0  8.0   */
/*                                |     with                                      |                                       */
/*                                |          recursive cnt(x) as (                |                                       */
/*                                |       select                                  | SAS                                   */
/*                                |          1                                    | PA PB       SUFFIX A B  C             */
/*                                |       union                                   |                                       */
/*                                |          all                                  | A  fail        6   1 8  1             */
/*                                |       select                                  | A  fail        1   5 7  9             */
/*                                |          x+1                                  | A  success     2   1 4  6             */
/*                                |       from                                    | A  success     1   1 3  5             */
/*                                |          cnt                                  | B  fail        5   1 7  9             */
/*                                |       where                                   | B  fail        1   4 6  8             */
/*                                |          x < 2                                | B  success     1   1 3  5             */
/*                                |       ),                                      | B  success     1   9 2  4             */
/*                                |     qs as                                     | C  fail        4   1 6  8             */
/*                                |         (select                               | C  fail        1   3 5  7             */
/*                                |           "select                             | C  success     9   1 2  4             */
/*                                |              pa                               | C  success     1   8 1  3             */
/*                                |             ,pb                               |                                       */
/*                                |             ,?  as suffix                     |                                       */
/*                                |             ,a?  as a                         |                                       */
/*                                |             ,b?  as b                         |                                       */
/*                                |             ,c?  as c                         |                                       */
/*                                |          from                                 |                                       */
/*                                |             have" as str)                     |                                       */
/*                                |     select                                    |                                       */
/*                                |       group_concat(replace(                   |                                       */
/*                                |          r.str                                |                                       */
/*                                |         ,"?"                                  |                                       */
/*                                |         ,cast(l.x as text))                   |                                       */
/*                                |         ," union all ") as cde                |                                       */
/*                                |     from                                      |                                       */
/*                                |         cnt as l left join qs as r            |                                       */
/*                                |     on 1=1                                    |                                       */
/*                                |    ''')                                       |                                       */
/*                                | txt=cde['cde'].astype(str).str.cat(sep=' ')   |                                       */
/*                                | print(txt)                                    |                                       */
/*                                | want=pdsql(txt)                               |                                       */
/*                                | print(want)                                   |                                       */
/*                                | ;;;;                                          |                                       */
/*                                | %utl_pyendx;                                  |                                       */
/*                                |                                               |                                       */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
  input Pa $ pb $
   a1 a2 b1 b2 c1 c2 ;
cards4;
A  success 1 2 3 4 5 6
A  fail    5 6 7 8 9 1
B  success 9 1 2 3 4 5
B  fail    4 5 6 7 8 9
C  success 8 9 1 2 3 4
C  fail    3 4 5 6 7 8
;;;;
run;quit;

/**************************************************************************************************************************/
/*  PA PB      A1 A2 B1 B2 C1 C2                                                                                          */
/*                                                                                                                        */
/*  A  success  1  2  3  4  5  6                                                                                          */
/*  A  fail     5  6  7  8  9  1                                                                                          */
/*  B  success  9  1  2  3  4  5                                                                                          */
/*  B  fail     4  5  6  7  8  9                                                                                          */
/*  C  success  8  9  1  2  3  4                                                                                          */
/**************************************************************************************************************************/


/*                             _
/ |  ___  __ _ ___   ___  __ _| |  _ __   ___    __ _ _ __ _ __ __ _ _   _ ___
| | / __|/ _` / __| / __|/ _` | | | `_ \ / _ \  / _` | `__| `__/ _` | | | / __|
| | \__ \ (_| \__ \ \__ \ (_| | | | | | | (_) || (_| | |  | | | (_| | |_| \__ \
|_| |___/\__,_|___/ |___/\__, |_| |_| |_|\___/  \__,_|_|  |_|  \__,_|\__, |___/
                            |_|                                      |___/
*/

proc sql;
 create
   table want as
 select
   pa
  ,pb
  ,1 as suffix
  ,a1 as a
  ,b1 as b
  ,c1 as c
 from
   sd1.have
 union
   all
 select
   pa
  ,pb
  ,1 as suffix
  ,a2 as a
  ,b2 as b
  ,c2 as c
 from
   sd1.have
 order
   by pa, pb, a
;quit;

/**************************************************************************************************************************/
/*  PA    PB         SUFFIX    A    B    C                                                                                */
/*                                                                                                                        */
/*  A     fail          1      5    7    9                                                                                */
/*  A     fail          1      6    8    1                                                                                */
/*  A     success       1      1    3    5                                                                                */
/*  A     success       1      2    4    6                                                                                */
/*  B     fail          1      4    6    8                                                                                */
/*  B     fail          1      5    7    9                                                                                */
/*  B     success       1      1    3    5                                                                                */
/*  B     success       1      9    2    4                                                                                */
/*  C     fail          1      3    5    7                                                                                */
/*  C     fail          1      4    6    8                                                                                */
/*  C     success       1      8    1    3                                                                                */
/*  C     success       1      9    2    4                                                                                */
/**************************************************************************************************************************/

/*___                              _
|___ \   ___  __ _ ___   ___  __ _| |   __ _ _ __ _ __ __ _ _   _ ___
  __) | / __|/ _` / __| / __|/ _` | |  / _` | `__| `__/ _` | | | / __|
 / __/  \__ \ (_| \__ \ \__ \ (_| | | | (_| | |  | | | (_| | |_| \__ \
|_____| |___/\__,_|___/ |___/\__, |_|  \__,_|_|  |_|  \__,_|\__, |___/
                                |_|                         |___/
*/

%array(_sfx,values=1 2);

proc sql;
  create
    table want as
    %do_over(_sfx,phrase=%str(
      select
        pa
       ,pb
       ,? as suffix
       ,a? as a
       ,b? as b
       ,c? as c
      from
        have), between=union all)
      order
        by pa, pb, suffix
;quit;

/**************************************************************************************************************************/
/*  PA    PB         SUFFIX    A    B    C                                                                                */
/*                                                                                                                        */
/*  A     fail          1      5    7    9                                                                                */
/*  A     fail          1      6    8    1                                                                                */
/*  A     success       1      1    3    5                                                                                */
/*  A     success       1      2    4    6                                                                                */
/*  B     fail          1      4    6    8                                                                                */
/*  B     fail          1      5    7    9                                                                                */
/*  B     success       1      1    3    5                                                                                */
/*  B     success       1      9    2    4                                                                                */
/*  C     fail          1      3    5    7                                                                                */
/*  C     fail          1      4    6    8                                                                                */
/*  C     success       1      8    1    3                                                                                */
/*  C     success       1      9    2    4                                                                                */
/**************************************************************************************************************************/


/*____                    _                                _
|___ /   _ __   ___  __ _| |  _ __ ___  ___ _   _ _ __ ___(_)_   _____
  |_ \  | `__| / __|/ _` | | | `__/ _ \/ __| | | | `__/ __| \ \ / / _ \
 ___) | | |    \__ \ (_| | | | | |  __/ (__| |_| | |  \__ \ |\ V /  __/
|____/  |_|    |___/\__, |_| |_|  \___|\___|\__,_|_|  |___/_| \_/ \___|
                       |_|
*/

%utl_rbeginx;
parmcards4;
library(haven)
library(sqldf)
source("c:/oto/fn_tosas9x.R")
options(sqldf.dll = "d:/dll/sqlean.dll")
have<-read_sas("d:/sd1/have.sas7bdat")
print(have)
cde<-sqldf('
 with
      recursive cnt(x) as (
   select
      1
   union
      all
   select
      x+1
   from
      cnt
   where
      x < 2
   ),
 qs as
     (select
       "select
          pa
         ,pb
         ,?  as suffix
         ,a?  as a
         ,b?  as b
         ,c?  as c
      from
         have" as str)
 select
   group_concat(replace(
      r.str
     ,"?"
     ,cast(l.x as text))
     ," union all ") as cde
 from
     cnt as l left join qs as r
 on 1=1
')
cde
want<-sqldf(cde$cde)
want
fn_tosas9x(
      inp    = want
     ,outlib ="d:/sd1/"
     ,outdsn ="want"
     )
;;;;
%utl_rendx;

proc print data=sd1.want;
run;quit;

/**************************************************************************************************************************/
/* R                            |  SAS                                                                                    */
/*    PA      PB suffix a b c   |  ROWNAMES    PA    PB         SUFFIX    A    B    C                                     */
/*                              |                                                                                         */
/* 1   A success      1 1 3 5   |      1       A     success       1      1    3    5                                     */
/* 2   A    fail      1 5 7 9   |      2       A     fail          1      5    7    9                                     */
/* 3   B success      1 9 2 4   |      3       B     success       1      9    2    4                                     */
/* 4   B    fail      1 4 6 8   |      4       B     fail          1      4    6    8                                     */
/* 5   C success      1 8 1 3   |      5       C     success       1      8    1    3                                     */
/* 6   C    fail      1 3 5 7   |      6       C     fail          1      3    5    7                                     */
/* 7   A success      2 2 4 6   |      7       A     success       2      2    4    6                                     */
/* 8   A    fail      2 6 8 1   |      8       A     fail          2      6    8    1                                     */
/* 9   B success      2 1 3 5   |      9       B     success       2      1    3    5                                     */
/* 10  B    fail      2 5 7 9   |     10       B     fail          2      5    7    9                                     */
/* 11  C success      2 9 2 4   |     11       C     success       2      9    2    4                                     */
/* 12  C    fail      2 4 6 8   |     12       C     fail          2      4    6    8                                     */
/**************************************************************************************************************************/

/*  _                 _   _                             _
| || |    _ __  _   _| |_| |__   ___  _ __    ___  __ _| |
| || |_  | `_ \| | | | __| `_ \ / _ \| `_ \  / __|/ _` | |
|__   _| | |_) | |_| | |_| | | | (_) | | | | \__ \ (_| | |
   |_|   | .__/ \__, |\__|_| |_|\___/|_| |_| |___/\__, |_|
         |_|    |___/                                |_|
*/

proc datasets lib=sd1 nolist nodetails;
 delete pywant;
run;quit;

%utl_pybeginx;
parmcards4;
import re
exec(open('c:/oto/fn_pythonx.py').read());
have,meta = ps.read_sas7bdat( \
  'd:/sd1/have.sas7bdat');
cde=pdsql('''
    with
         recursive cnt(x) as (
      select
         1
      union
         all
      select
         x+1
      from
         cnt
      where
         x < 2
      ),
    qs as
        (select
          "select
             pa
            ,pb
            ,?  as suffix
            ,a?  as a
            ,b?  as b
            ,c?  as c
         from
            have" as str)
    select
      group_concat(replace(
         r.str
        ,"?"
        ,cast(l.x as text))
        ," union all ") as cde
    from
        cnt as l left join qs as r
    on 1=1
   ''')
txt=cde['cde'].astype(str).str.cat(sep=' ')
print(txt)
want=pdsql(txt)
print(want)
fn_tosas9x(want,outlib='d:/sd1/',outdsn='pywant',timeest=3);
;;;;
%utl_pyendx;

proc print data=sd1.pywant;
run;quit;

/**************************************************************************************************************************/
/* Python                                  |  SAS                                                                         */
/*    PA       PB  suffix    a    b    c   |    PA    PB         SUFFIX    A    B    C                                    */
/*                                         |                                                                              */
/* 0   A  success       1  1.0  3.0  5.0   |    A     success       1      1    3    5                                    */
/* 1   A     fail       1  5.0  7.0  9.0   |    A     fail          1      5    7    9                                    */
/* 2   B  success       1  9.0  2.0  4.0   |    B     success       1      9    2    4                                    */
/* 3   B     fail       1  4.0  6.0  8.0   |    B     fail          1      4    6    8                                    */
/* 4   C  success       1  8.0  1.0  3.0   |    C     success       1      8    1    3                                    */
/* 5   C     fail       1  3.0  5.0  7.0   |    C     fail          1      3    5    7                                    */
/* 6   A  success       2  2.0  4.0  6.0   |    A     success       2      2    4    6                                    */
/* 7   A     fail       2  6.0  8.0  1.0   |    A     fail          2      6    8    1                                    */
/* 8   B  success       2  1.0  3.0  5.0   |    B     success       2      1    3    5                                    */
/* 9   B     fail       2  5.0  7.0  9.0   |    B     fail          2      5    7    9                                    */
/* 10  C  success       2  9.0  2.0  4.0   |    C     success       2      9    2    4                                    */
/* 11  C     fail       2  4.0  6.0  8.0   |    C     fail          2      4    6    8                                    */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
