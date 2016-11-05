Attribute VB_Name = "Declaration"
Option Explicit

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'---------------------------------------------------------------
Public strCN           As New A1
Public U_Id            As String
'--------------------------------------------------------------

Public Rpt_Nm           As String       ' > Report viewer

Public Rpt_Desig        As String
Public Rpt_EmpID       As String
Public Rpt_Unit         As String
Public Rpt_Cost         As String

Public Rpt_From         As Date
Public Rpt_To           As Date
Public Rpt_Date         As String
Public Rpt_Month        As String
Public Rpt_Year         As String
Public Rpt_Fiscal_Yr    As String
Global GetYearfromSalary
Global ComboValue_Dept
Global Emp_ID_Value
Global Emp_ID_Value_ForLoan
Global EnddateforReport
Global BeginDateForReport
Global Emp_IDforLeave
Global BeginDateOfIncremnt
Global EndDateOfIncremnt
Global GetMonthOftheYear
Global StatusofEmployee
Global EmployeeName
Global EmpIDForTowhom
Global BEGINYEARFORWHOM
Global ENDDATEFORWHOM
Global BonusPreparationStatus As Integer
Global DEPARMENTNAMEFORTPT As String
Global SEXFORREPORT As Integer
Global DESIGNATIONFORRPT As String
Global CheckStatusofEmployee As Integer
Global ReportStatusofEmployee As Integer
Global GetSalaryPreparationYaer As Integer
Global DesignationOfEmp As String
Global DepartemntOfEmp As String
Global JobType As String
Global ReportTracker As String
Global DateofRetirement As String
Global DatofJoin As String
Global GetFromMonthtoWhom  As String
Global GetToMonthtoWhom As String

Global sBankCode As String
Global sAccountType As String
Global sAccountNo As String
Global sBankName As String
Global sAccountTypeName As String
Global sSourceId As String
Global sSourceName As String
Global sPurposeId As String
Global sPurposeName As String
Global sGender As String
Public Const organizationInfo As String = "IT Division, DNMIH"

Public Const Id_Len As Integer = 4      ' Max length of
Global localSalaryType As String
Global twoPage As Integer
Global underDepartmentorNot As Integer

Global currentOption As String
Global currentDept As String
Global currentFormat As Integer
Global paramMonth As String
Global paramYear As String
Global paramDepartment As String
Public UserRole As String
Public pScale As String
Public Const yRangeForSDA As Integer = 2
Public Const jobTypePermanent As String = "003"
Public Const allowanceLawDate As String = "05-oct-2010"
