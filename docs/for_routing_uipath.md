# CA01 element IDs handed to the UiPath side

Three SAP GUI element IDs for the CA01 initial screen, recorded for the UiPath routing
project. **Keep this file:** `ROUTING\docs\uipath\ui_elements.md` section 2 cites it by name
as the source it verified the UiPath `create_routing_*.xaml` flows against (all three
matched; UiPath double-clicks copy-from where Python would use one `press()`).

CREATE ROUTING CA01:
material field: "<sap id='usr/ctxt[0]' />"
type into plant 1901: "<sap id='usr/ctxt[1]' />"
click copy from: "<sap id='tbar[1]/btn[5]' />"

A `usr/...` id here is the same element SAP GUI scripting addresses as
`wnd[0]/usr/...`. CA01 is not implemented in Python yet - plan in
[`TODO_bug.md`](TODO_bug.md) section E.
