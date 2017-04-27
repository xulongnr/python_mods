@echo off
set MyClassPath=autoTest\WEB-INF\classes
set MyMainClass=sgcc.nds.jdbc.testmain.TestMain
set TestCasePrefix=sgcc.nds.jdbc.testcase

::TestPreparedStatement
::TestStatement
::TestCallableStatement
::TestBlob
::TestClob
::TestDriver
::TestConnection
::TestSavepoint
::TestResultSet
::TestResultSetMetaData
::TestParameterMetaData
::statementTest.TestMemValve

set TestCaseList=TestPreparedStatement TestStatement TestCallableStatement TestBlob TestClob TestDriver TestConnection TestSavepoint TestResultSet TestResultSetMetaData TestParameterMetaData
set TestCaseList=TestResultSet
set TestCaseFunction=test_getCursorName

(for %%t in (%TestCaseList%) do (
   echo Testing ========= %%t::%TestCaseFunction% ==========
   java -classpath %MyClassPath% %MyMainClass% :%TestCasePrefix%.%%t:%TestCaseFunction%
))

::java -classpath %MyClassPath% %MyMainClass% :%TestCasePrefix%.TestPreparedStatement:* :%TestCasePrefix%.TestStatement:*

set str_time_first_bit="%time:~0,1%"
if %str_time_first_bit%==" " (
    set str_date_time=%date:~0,4%%date:~5,2%%date:~8,2%_0%time:~1,1%%time:~3,2%%time:~6,2%
)else (
    set str_date_time=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%
)
set foldername=%str_date_time%
set folder_prefix=autoTest\log\
set folder_name=%folder_prefix%%foldername%
mkdir %folder_name%
move %folder_prefix%*.log %folder_name%
if /i "%TestCaseFunction%" == "*" (
	cd %folder_name%
	python ..\..\..\testresult2xlsx.py
	cd ..\..\..
)

@echo on
