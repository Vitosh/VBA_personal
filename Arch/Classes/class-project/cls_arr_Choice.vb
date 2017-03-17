Option Explicit

Private p_investor                  As String
Private p_region                    As String
Private p_standort                  As String
Private p_project                   As String
Private p_ba_number                 As Long
Private p_global                    As Boolean

Public Property Get Investor() As String
    Investor = p_investor
End Property

Public Property Let Investor(str_investor_type As String)
    p_investor = str_investor_type
End Property

Public Property Get Region() As String
    Region = p_region
End Property

Public Property Let Region(str_region As String)
    p_region = str_region
    p_standort = IIf(str_region = "Wien", "Austria", "Germany")
End Property

Public Property Get Standort()
    Standort = p_standort
End Property

Public Property Get Project() As String
    Project = p_project
End Property

Public Property Let Project(str_project As String)
    p_project = str_project
End Property

Public Property Get BAnumber() As Long
    BAnumber = p_ba_number
End Property

Public Property Let BAnumber(l_ba_number As Long)
    p_ba_number = l_ba_number
End Property

Public Property Let GlobalProject(b_is_global As Boolean)
    p_global = b_is_global
End Property

Public Property Get GlobalProject() As Boolean
    GlobalProject = p_global
End Property

Public Property Get GewerbeGlobal() As Boolean
    
    If GlobalProject And Project = type_string_project(enum_project.project_gewerbe) Then
        GewerbeGlobal = True
    Else
        GewerbeGlobal = False
    End If
    
End Property

