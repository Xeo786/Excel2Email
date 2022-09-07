; Excel2Email method is simple example about, how to Appened a MS Excel range onto MS Outlook Email as (MS words) Table
; it supports both kind of Emails (New Email created with default Signature / without any Signatures
; https://github.com/Xeo786/Excel2Email
; By Xeo786
Excel2Email(table,Mail,Mailbody)
{
	t_rows := table.Rows.count
	t_cols := table.Columns.count
	myInspector := Mail.GetInspector
	Doc := myInspector.WordEditor
	wRange := Doc.Range(0, Doc.Characters.Count)
	wRange.InsertBefore("`n")
	oltable := Doc.Tables.add(Doc.Range(1, 1),t_rows,t_cols)
	r := 0
	loop, % t_rows
	{	
		++r
		c := 0
		loop, % t_cols
		{
			++c
			oltable.Cell(r,c).range.Text := table.cells(r,c).text
		}
	}
	oltable.ApplyStyleRowBands := true
	oltable.Borders.InsideLineStyle := 1
	oltable.Borders.OutsideLineStyle := 1
	oltable.AutoFitBehavior(1)
	wRange.InsertBefore(Mailbody "`n")
}
