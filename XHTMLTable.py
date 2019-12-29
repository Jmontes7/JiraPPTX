class XHTMLTable:
    """Class to create a XHTML Tables in Confluence """

    def __init__(self):
        self

    def InsertIssue(self, Issue ):
        """Insert a new link to JIRa Issue"""
        Text='<div class="content-wrapper"> \
             <p><ac:structured-macro ac:name="jira" ac:schema-version="1"  \
             ac:macro-id="2df1f18a-2f01-4b73-aa95-a429f196bfc4"> \
             <ac:parameter ac:name="server">VBPS JIRA</ac:parameter> \
             <ac:parameter ac:name="columns"> \
             key,summary,type,created,updated,due,assignee,reporter,priority,status,resolution \
             </ac:parameter><ac:parameter ac:name="serverId">97e8ad75-aef4-3f1a-af05-bfc704ed125c \
             </ac:parameter><ac:parameter ac:name="key">' + Issue + '</ac:parameter></ac:structured-macro></p></div>'
        return Text;


    def InsertColName (self, ColText):
        """Insert an Sprint Name with center Aligment and Bold font"""
        TextHTML = ColText.replace("&", "&amp;")
        Text= '<h2 style="text-align: center;"><strong>' + TextHTML + '</strong></h2>'
        return Text;

    def InsertRowName (self, RowText):
        """Insert a Team Name"""
        TextHTML = RowText.replace("&", "&amp;")
        Text= '<h2><strong>' + TextHTML + '</strong></h2>'
        return Text;

    def CreateTable(self):
        """Creates the initial XHTML to create a table - Header"""
        Text = '<table class="relative-table wrapped" style="width: 93.05%;"><tbody>'
        return Text;

    def CloseTable(self):
        """Creates the XHTML code to close a table"""
        Text = '</tbody></table>'
        return Text;

    def CreateRow(self):
        """Creates a new Row in XHTML"""
        Text = str("<tr>")
        return Text;

    def CloseRow(self):
        """Close the Row"""
        Text = str("</tr>")
        return Text;

    def CreateColumn(self):
        """Creates a new Column in XHTML"""
        Text = '<th class="highlight-red" title="Background colour : Red" data-highlight-colour="red">'
        return Text;

    def CloseColumn(self):
        """Close the column"""
        Text = '</th>'
        return Text;

    def CreateCell(self):
        """Creates a new Cell in XHTML"""
        Text = str("<td>")
        return Text;

    def CreateCenterCell(self):
        """Creates a new Cell in XHTML"""
        Text = str('<td style="text-align: center;">')
        return Text;

    def CreateMergedCell(self,rowspan):
        """Creates a new Cell in XHTML"""

        Text = str('<td rowspan="' + str(rowspan) + '">')
        return Text;

    def CloseCell(self):
        """Closes the Cell"""
        Text = str("</td>")
        return Text;

