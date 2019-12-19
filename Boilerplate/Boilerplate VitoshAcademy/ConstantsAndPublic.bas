Attribute VB_Name = "ConstantsAndPublic"
Option Explicit
Option Private Module

Public Const SET_IN_PRODUCTION = True
Public Const WORKSHEET_UNPROTECT_PASSWORD = "shouldistayorshouldigo"    'I am never using this password anywhere, do not bother ;)
Public Const ADMINS = "vitosh:vitos"
Public Const CON_STR_APP_NAME = "Boilerplate VitoshAcademy"
Public Const CON_STR_INSTANCES_LOG = "More then one Workbook is opened in this Excel instance."
Public Const CON_STR_1904 = "You are using 1904 date system. This is probably* not what you need."

'Public variables are a bad practice and should be avoided in general...
Public PUB_STR_ERROR_REPORT As String
