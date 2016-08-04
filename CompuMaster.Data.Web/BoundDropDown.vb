Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls


<Assembly: TagPrefix("CompuMaster.Data.Web", "CMDataWeb")> 
Namespace CompuMaster.Data.Web
	<DefaultProperty("Text"), ToolboxData("<{0}:BoundDropDown runat=server></{0}:BoundDropDown>")> _
	Public Class BoundDropDown
		Inherits DropDownList

		Public BoundField As String

	End Class
End Namespace
