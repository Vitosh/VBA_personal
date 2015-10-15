Option Explicit

Sub Load_Data_To_Object()

    Set my_choice = New cls_arrChoice
    
    my_choice.Investor = type_string_investor(enum_investors.inv_Private)
    my_choice.Region = type_string_standort(enum_standort.standort_Vienna)
    my_choice.Project = type_string_project(enum_project.project_gewerbe)
    my_choice.BAnumber = enum_BA.BA_10
    my_choice.GlobalProject = True
    
End Sub

Sub Display_Data_From_Object()

    Debug.Print my_choice.Investor
    Debug.Print my_choice.Standort
    Debug.Print my_choice.Region
    Debug.Print my_choice.Project
    Debug.Print my_choice.BAnumber
    Debug.Print my_choice.GlobalProject
    Debug.Print my_choice.GewerbeGlobal
    
End Sub
