import re

from Connect import Connect
from XHTMLTable import XHTMLTable


class Tools:
    """Class to work with JIRA Data"""

    global config
    global jira
    global confluence

    # Create Table Class object
    Connect = Connect()

    # Connect
    confluence = Connect.Confluence()
    jira = Connect.Jira()
    config = Connect.GetConfig()

    def __init__(self):
        pass

    def SprintMatrix(self):
        """Method that creates an array with the Sprints"""

        # Initialize Variables
        ArraySprint = []

        # Prepare JIRA Query
        JIRAQuery = config['INI'].get('FeatureJQL') + ' ORDER BY "Sprint" ASC'

        # Execute Feature Query in JIRA
        Issues = jira.search_issues(JIRAQuery, startAt=0, maxResults=False)

        for issue in Issues:
            found = False
            for i in range(len(ArraySprint)):
                if ArraySprint[i] == str(self.FindSprint(issue)):
                    found = True
            if found == False:
                ArraySprint.append(str(self.FindSprint(issue)))

        return ArraySprint;

    def TeamMatrix(self):
        """Method that creates an array with the teams"""

        # Initialize Variables
        ArrayTeam = []

        # Prepare JIRA Query
        JIRAQuery = config['INI'].get('FeatureJQL') + ' ORDER BY "Assigned Team" ASC'

        # Execute Feature Query in JIRA
        Issues = jira.search_issues(JIRAQuery, startAt=0, maxResults=False)

        for issue in Issues:
            # Find Team
            found = False
            for i in range(len(ArrayTeam)):
                if ArrayTeam[i] == str(issue.fields.customfield_11018):
                    found = True

            if found == False:
                ArrayTeam.append(str(issue.fields.customfield_11018))

        return ArrayTeam;

    def AssigneeMatrix(self):
        """Method that creates an array with the Assignees"""

        # Initialize Variables
        ArrayAssignee = []

        # Prepare JIRA Query
        JIRAQuery = config['INI'].get('StoryJQL') + ' ORDER BY "Assignee" ASC'

        # Execute Story Query in JIRA
        Issues = jira.search_issues(JIRAQuery, startAt=0, maxResults=False)

        for issue in Issues:
            # Find Assignee
            found = False
            for i in range(len(ArrayAssignee)):
               try:
                 if ArrayAssignee[i] == str(issue.fields.assignee.displayName):
                        found = True
               except:
                   pass

            if found == False:
                try:
                   ArrayAssignee.append(str(issue.fields.assignee.displayName))
                except:
                    pass

        return ArrayAssignee;

    def SUMPoints(self, IssueId):
        # Stories per Feature
        StoryQuery = config['INI'].get('StoryJQL')  + " and 'Epic Link'=" + IssueId
        storySP = 0

        # Execute Story Query in JIRA
        Stories = jira.search_issues(StoryQuery, startAt=0, maxResults=False)

        for issues in Stories:

            if issues.fields.customfield_10002:
                storySP = storySP + issues.fields.customfield_10002

        return storySP;

    def FindSprint(self, issue):
        # Find Feature Sprint
        try:
            for sprint in issue.fields.customfield_10004:
                sprint_startDate = re.findall(r"startDate=[^,]*", str(issue.fields.customfield_10004[0]))
                sprint_endDate = re.findall(r"endDate=[^,]*", str(issue.fields.customfield_10004[0]))
                sprint_completeDate = re.findall(r"completeDate=[^,]*", str(issue.fields.customfield_10004[0]))
                sprint_state = re.findall(r"state=[^,]*", str(issue.fields.customfield_10004[0]))
                sprint_Id = re.findall(r"sequence=[^,]*", str(issue.fields.customfield_10004[0]))
                sprint_name = re.findall(r"name=[^,]*", str(issue.fields.customfield_10004[0]))
                Sprint_text = str(sprint_name)
                Start_Sprint = Sprint_text.find("Sprint")
                Total_SprintLen = len(Sprint_text)
                # Get the right name for the sprint
                FeatSprint = Sprint_text[Start_Sprint:Total_SprintLen - 2]
        except:
            FeatSprint = "N/A"

        return  FeatSprint;

    def WriteConfluence(self,PAGE, Prefix,  HTMLBody):
        # Check if the page exist
        PageID = confluence.get_page_id(config[PAGE].get('space'), Prefix + " " + config[PAGE].get('title'))

        # If exist Delete the page
        if PageID != "None":
            confluence.remove_page(PageID, status=None)

        status = confluence.create_page(
            space=config[PAGE].get('space'),
            title=Prefix + " " + config[PAGE].get('title'),
            parent_id=config[PAGE].get('parent_id'),
            body=HTMLBody)
        print(status)
        return;

    def LastComment(self, issue):

        # Review comments
        comment = jira.comments(issue.key)

        #find the last comment
        a = int(0)
        for c in comment:
            if int(c.id) > a:
                a = int(c.id)

        # Print Comment
        try:
           # ELiminate Wrong Characters
           CommentText = jira.comment(issue.key, a).body
           TextHTML = CommentText.replace("&", "&amp;")
        except:
           TextHTML = ""

        return TextHTML;

    def GetFixVersion(self, issue):

        # GetVersion
        try:
            GetVersion = str(issue.fields.fixVersion)
        except:
            GetVersion = ""

        return GetVersion;


    # Find Feature dependencies

    def InsertDependencies(self, HTMLBody, issue):

        #Create object
        XHMTL = XHTMLTable()

        HTMLBody = HTMLBody + XHMTL.CreateCell()

        OutCount = 0
        InCount = 0

        for link in issue.fields.issuelinks:
            if hasattr(link, "outwardIssue"):
                if OutCount == 0:
                    HTMLBody = HTMLBody + "<p><strong>dependents on</strong></p>"
                    OutCount += 1
                outwardIssue = link.outwardIssue
                HTMLBody = HTMLBody + XHMTL.InsertIssue(outwardIssue.key)
            if hasattr(link, "inwardIssue"):
                if InCount == 0:
                    HTMLBody = HTMLBody + "<p><strong>is depended on by</strong></p>"
                    InCount += 1
                inwardIssue = link.inwardIssue
                HTMLBody = HTMLBody + XHMTL.InsertIssue(inwardIssue.key)

        HTMLBody = HTMLBody + XHMTL.CloseCell()

        return HTMLBody;

    def CleanTxt (self, Texto):

        Texto = Texto.replace("&", "&amp;")

        return Texto;

