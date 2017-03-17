ToDo in a VBA project (Tasks for a boilerplate) :

> 	Make OnStart and OnEnd modules
>	Make if [set_in_production] then on error goto Main_error
>	Play with the status bar
>	Show a vbmodeless form while the macro is running
> 	Find a quick macro to lock and unlock the project
>	On start of the file:
		> 	Lock it
		>	Hide Not needed sheets
		>	Lock scroll
		>	Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", false)"
>	On the end of the file:
		>	Save and release all possible forbidden things (check Workbook_BeforeClose)
>	Disable Workbook_NewSheet
>	Make a quick unlock and view all option just for you
>	Disable copy and paste and F11
