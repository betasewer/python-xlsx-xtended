"""
ssml:sheetViews = ssml:CT_SheetViews
Sequence [1..1]
    ssml:sheetView [1..*]    Worksheet View
    ssml:extLst [0..1]    Future Feature Storage Area

ssml:sheetView = ssml:CT_SheetView
Sequence [1..1]
    ssml:pane [0..1]    View Pane
    ssml:selection [0..4]    Selection
    ssml:pivotSelection [0..4]    PivotTable Selection
    ssml:extLst [0..1]    Future Feature Storage Area

ssml:pane = ssml:CT_Pane
Attribute
    xSplit	[0..1]	xsd:double	Horizontal Split Position	Default value is "0".
    ySplit	[0..1]	xsd:double	Vertical Split Position	Default value is "0".
    topLeftCell	[0..1]	ssml:ST_CellRef	Top Left Visible Cell	
    activePane	[0..1]	ssml:ST_Pane	Active Pane	Default value is "topLeft".
    state	[0..1]	ssml:ST_PaneState	Split State	Default value is "split". 

ssml:selection = ssml:CT_Selection
Atrribute
    pane	[0..1]	ssml:ST_Pane	Pane	Default value is "topLeft".
    activeCell	[0..1]	ssml:ST_CellRef	Active Cell Location	
    activeCellId	[0..1]	xsd:unsignedInt	Active Cell Index	Default value is "0".
    sqref	[0..1]	ssml:ST_Sqref	Sequence of References	Default value is "A1". 

ssml:pivotSelection = ssml:CT_PivotSelection
Atrribute
    pane	[0..1]	ssml:ST_Pane	Pane	Default value is "topLeft".
    showHeader	[0..1]	xsd:boolean	Show Header	Default value is "false".
    label	[0..1]	xsd:boolean	Label	Default value is "false".
    data	[0..1]	xsd:boolean	Data Selection	Default value is "false".
    extendable	[0..1]	xsd:boolean	Extendable	Default value is "false".
    count	[0..1]	xsd:unsignedInt	Selection Count	Default value is "0".
    axis	[0..1]	ssml:ST_Axis	Axis	
    dimension	[0..1]	xsd:unsignedInt	Dimension	Default value is "0".
    start	[0..1]	xsd:unsignedInt	Start	Default value is "0".
    min	[0..1]	xsd:unsignedInt	Minimum	Default value is "0".
    max	[0..1]	xsd:unsignedInt	Maximum	Default value is "0".
    activeRow	[0..1]	xsd:unsignedInt	Active Row	Default value is "0".
    activeCol	[0..1]	xsd:unsignedInt	Active Column	Default value is "0".
    previousRow	[0..1]	xsd:unsignedInt	Previous Row	Default value is "0".
    previousCol	[0..1]	xsd:unsignedInt	Previous Column Selection	Default value is "0".
    click	[0..1]	xsd:unsignedInt	Click Count	Default value is "0".
    r:id	[0..1]	r:ST_RelationshipId	Relationship Id

"""