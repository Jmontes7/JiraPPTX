from Connect import Connect
from Tools import Tools
from XHTMLTable import XHTMLTable

class ConfReports:
    global config
    global jira
    global confluence
    global XHMTL
    global Tools
    global Text_Label

    # Create Table Class object
    XHMTL = XHTMLTable()
    Connect = Connect()
    Tools = Tools()

    # Fill parameters
    config = Connect.GetConfig()
    jira = Connect.Jira()

    # Connect
    confluence = Connect.Confluence()
    jira = Connect.Jira()
    config = Connect.GetConfig()

    #Fill Arrays
    ArraySprint = Tools.SprintMatrix()
    ArrayTeam = Tools.TeamMatrix()
    ArrayAssignee = Tools.AssigneeMatrix()

    def __init__(self):
        pass

    # ********************************
    # PI Programme Board Report
    # ********************************

    def PIBoard(self):
        # Create Table
        HTMLBody = XHMTL.CreateTable()

        # Create Columns
        HTMLBody = HTMLBody + XHMTL.CreateRow()

        HTMLBody = HTMLBody + XHMTL.CreateColumn()
        HTMLBody = HTMLBody + XHMTL.InsertColName("Teams")
        HTMLBody = HTMLBody + XHMTL.CloseColumn()

        # Add Sprint Name in Column header
        for i in range(len(self.ArraySprint)):
            HTMLBody = HTMLBody + XHMTL.CreateColumn()
            HTMLBody = HTMLBody + XHMTL.InsertColName(self.ArraySprint[i])
            HTMLBody = HTMLBody + XHMTL.CloseColumn()
        HTMLBody = HTMLBody + XHMTL.CloseRow()

        # For each team
        for i in range(len(self.ArrayTeam)):
            HTMLBody = HTMLBody + XHMTL.CreateRow()
            HTMLBody = HTMLBody + XHMTL.CreateCell()
            HTMLBody = HTMLBody + XHMTL.InsertRowName(self.ArrayTeam[i])
            HTMLBody = HTMLBody + XHMTL.CloseCell()

            # for each sprint
            for j in range(len(self.ArraySprint)):
                HTMLBody = HTMLBody + XHMTL.CreateCell()

                # Create JIRA Query
                if self.ArraySprint[j] != "N/A":
                    Query = config['INI'].get('FeatureJQL') + ' AND Sprint ="' + self.ArraySprint[j] + '"'
                else:
                    Query = config['INI'].get('FeatureJQL') + ' AND Sprint is EMPTY'

                if self.ArrayTeam[i] != "None":
                    Query = Query + ' AND "Assigned Team"="' + self.ArrayTeam[i] + '"'
                else:
                    Query = Query + ' AND "Assigned Team" is EMPTY'

                print(Query)
                # Execute Feature Query in JIRA
                Issues = jira.search_issues(Query, startAt=0, maxResults=False)
                for issue in Issues:
                    HTMLBody = HTMLBody + XHMTL.InsertIssue(issue.key)
                    print('Inserted Feature ' + issue.key)
                HTMLBody = HTMLBody + XHMTL.CloseCell()
            HTMLBody = HTMLBody + XHMTL.CloseRow()

        # CloseTable
        HTMLBody = HTMLBody + XHMTL.CloseTable()

        # Write Page in Confluence
        Tools.WriteConfluence('PI-PAGE', "", HTMLBody)
        Text_Label = "Report Completed"
        return Text_Label;

    # ********************************
    # Team Board Report
    # ********************************
    def TeamBoard(self):

        #Check One Team Parameter
        #if config['REPORT'].get('OneTeam') == False:
        Teams = len(self.ArrayTeam)
        #else:
        #  Teams = 1

        # Create a New Page for each Team
        for i in range(Teams):
            # ********************************
            # Create Table
            # ********************************

            HTMLBody = XHMTL.CreateTable()

            # ********************************
            # Create Columns
            # ********************************

            ColName = ['Features', 'Dependencies', 'Stories', 'Assignee', 'Sprint', 'Story Points', 'Comments',
                       'Priority']

            # New Row
            HTMLBody = HTMLBody + XHMTL.CreateRow()

            for Col in range(len(ColName)):
                HTMLBody = HTMLBody + XHMTL.CreateColumn()
                HTMLBody = HTMLBody + XHMTL.InsertColName(ColName[Col])
                HTMLBody = HTMLBody + XHMTL.CloseColumn()

            HTMLBody = HTMLBody + XHMTL.CloseRow()

            # End Columns

            # ********************************
            # Add Rows
            # ********************************

            # Query JIRA
            #if config['REPORT'].get('OneTeam') == False:
            if self.ArrayTeam[i] != "None":
                Query = config['INI'].get('FeatureJQL') + ' AND "Assigned Team"="' + self.ArrayTeam[i] + '"'
            else:
                Query = config['INI'].get('FeatureJQL') + ' AND "Assigned Team" is EMPTY'
            #else:
            #    Query = config['INI'].get('FeatureJQL') + ' AND resolution = Unresolved'

            # Add Order By
            Query = Query + " " + config['INI'].get('FeatORDER')

            # ********************************
            # Insert Feature
            # ********************************

            print(Query)
            # Execute Feature Query in JIRA
            Issues = jira.search_issues(Query, startAt=0, maxResults=False)

            for issue in Issues:
                HTMLBody = HTMLBody + XHMTL.CreateRow()

                HTMLBody = HTMLBody + XHMTL.CreateCell()
                HTMLBody = HTMLBody + XHMTL.InsertIssue(issue.key)
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Dependencies
                # ********************************

                HTMLBody = Tools.InsertDependencies(HTMLBody, issue)

                # ********************************
                # Insert Stories
                # ********************************

                # Stories per Feature
                #if config['REPORT'].get('OneTeam') == False:
                StoryQuery = config['INI'].get('StoryJQL') + " and 'Epic Link'=" + issue.key
                StoryQuery = StoryQuery + " " + config['INI'].get('StoryORDER')
                #else:
                    #StoryQuery = config['INI'].get('StoryJQL') + " AND resolution = Unresolved AND 'Epic Link'=" + issue.key
                    #StoryQuery = StoryQuery + " " + config['INI'].get('StoryORDER')

                # Calculate Feature Story Points
                storySP = Tools.SUMPoints(issue.key)

                # Execute Story Query in JIRA
                Stories = jira.search_issues(StoryQuery, startAt=0, maxResults=False)

                # Create Cell
                HTMLBody = HTMLBody + XHMTL.CreateCell()
                for issues in Stories:
                    HTMLBody = HTMLBody + XHMTL.InsertIssue(issues.key)

                # Close Story Cell
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Assignee
                # ********************************

                # Create Cell
                HTMLBody = HTMLBody + XHMTL.CreateCell()
                if config['TEAM-PAGE'].get('Consolidate') == "True":
                    HTMLBody = HTMLBody + str(issue.fields.assignee.displayName)
                else:
                    for issues in Stories:
                        # Find Story Assignee
                        try:
                            HTMLBody = HTMLBody + "<p>" + str(issues.fields.assignee.displayName) + "</p>"
                        except:
                            HTMLBody = HTMLBody + "<p>N/A</p>"

                # Close Story Cell
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Sprint
                # ********************************
                HTMLBody = HTMLBody + XHMTL.CreateCell()

                if config['REPORT'].get('Consolidate') == "True":
                    # Find Sprint
                    HTMLBody = HTMLBody + Tools.FindSprint(issue)
                else:
                    for issues in Stories:
                        # Find Sprint
                        HTMLBody = HTMLBody + "<p>" + Tools.FindSprint(issues) + "</p>"

                # Close Sprint Cell
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Story Points
                # ********************************
                # Create Story Points Cell
                HTMLBody = HTMLBody + XHMTL.CreateCell()
                if config['REPORT'].get('Consolidate') == "True":
                    HTMLBody = HTMLBody + str(storySP)
                else:
                    for issues in Stories:
                        if issues.fields.customfield_10002:
                            HTMLBody = HTMLBody + "<p>" + str(issues.fields.customfield_10002) + "</p>"

                # Close Story Points Cell
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Comments
                # ********************************
                HTMLBody = HTMLBody + XHMTL.CreateCell()

                if config['REPORT'].get('Comments') == "True":
                    # Get the last comment of an Issue
                    HTMLBody = HTMLBody + Tools.LastComment(issue)
                else:
                    HTMLBody = HTMLBody + " "

                HTMLBody = HTMLBody + XHMTL.CloseCell()

                # ********************************
                # Insert Priority
                # ********************************
                HTMLBody = HTMLBody + XHMTL.CreateCell()
                HTMLBody = HTMLBody + str(issue.fields.priority)
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                HTMLBody = HTMLBody + XHMTL.CloseRow()

                print('Inserted Feature ' + issue.key)

            # CloseTable
            HTMLBody = HTMLBody + XHMTL.CloseTable()

            # Write Page in Confluence
            #if config['REPORT'].get('OneTeam') == False:
            Tools.WriteConfluence('TEAM-PAGE', self.ArrayTeam[i], HTMLBody)
            #else:
            #    Tools.WriteConfluence('TEAM-PAGE', "One Team", HTMLBody)


        return;

    # ********************************
    # Assignee Report
    # ********************************
    def AssigneeBoard(self):

        # Create a New Page for each Assignee
        for i in range(len(self.ArrayAssignee)):

            # Initiate varibales
            CountTables = ["Features", "Stories"]
            ColName = ['Dependencies', 'Sprint', 'Story Points', 'Comments', 'Priority']
            HTMLBody = ""

            # Loop run 2 times (Features and Stories)
            for x in CountTables:

                # ********************************
                # Create Table
                # ********************************

                HTMLBody = HTMLBody + XHMTL.CreateTable()

                # ********************************
                # Create Columns
                # ********************************

                # Initial Row
                HTMLBody = HTMLBody + XHMTL.CreateRow()

                # Create Column for Feature or Story
                HTMLBody = HTMLBody + XHMTL.CreateColumn()
                HTMLBody = HTMLBody + XHMTL.InsertColName(x)
                HTMLBody = HTMLBody + XHMTL.CloseColumn()

                for Col in range(len(ColName)):
                    HTMLBody = HTMLBody + XHMTL.CreateColumn()
                    HTMLBody = HTMLBody + XHMTL.InsertColName(ColName[Col])
                    HTMLBody = HTMLBody + XHMTL.CloseColumn()

                # Close Initial Row
                HTMLBody = HTMLBody + XHMTL.CloseRow()

                # End Columns

                # ********************************
                # Add Rows
                # ********************************

                # Prepare Query JIRA
                if x == "Features":
                    if self.ArrayAssignee[i] != "None":
                        Query = config['INI'].get('FeatureJQL') + ' AND "Assignee"="' + self.ArrayAssignee[i] + '"'
                    else:
                        Query = config['INI'].get('FeatureJQL') + ' AND "Assigned Team" is EMPTY'
                else:
                    if self.ArrayAssignee[i] != "None":
                        Query = config['INI'].get('StoryJQL') + ' AND "Assignee"="' + self.ArrayAssignee[i] + '"'
                    else:
                        Query = config['INI'].get('StoryJQL') + ' AND "Assigned Team" is EMPTY'

                # Add Order By
                if x == "Features":
                    Query = Query + " " + config['INI'].get('FeatORDER')
                else:
                    Query = Query + " " + config['INI'].get('StoryORDER')

                # ********************************
                # Insert New Issue
                # ********************************
                print(Query)
                # Execute Query in JIRA
                Issues = jira.search_issues(Query, startAt=0, maxResults=False)
                for issue in Issues:
                    # New Row
                    HTMLBody = HTMLBody + XHMTL.CreateRow()

                    # ********************************
                    # Insert Issue ID
                    # ********************************
                    HTMLBody = HTMLBody + XHMTL.CreateCell()
                    HTMLBody = HTMLBody + XHMTL.InsertIssue(issue.key)
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # ********************************
                    # Insert Dependencies
                    # ********************************

                    HTMLBody = Tools.InsertDependencies(HTMLBody, issue)

                    # ********************************
                    # Insert Sprint
                    # ********************************

                    HTMLBody = HTMLBody + XHMTL.CreateCell()
                    HTMLBody = HTMLBody + Tools.FindSprint(issue)
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # ********************************
                    # Insert Story Points
                    # ********************************
                    HTMLBody = HTMLBody + XHMTL.CreateCell()
                    # Create Story Points Cell
                    if x == "Features":
                        HTMLBody = HTMLBody + str(Tools.SUMPoints(issue.key))
                    else:
                        HTMLBody = HTMLBody + str(issue.fields.customfield_10002)
                    # Close Story Points Cell
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # ********************************
                    # Insert Comments
                    # ********************************
                    HTMLBody = HTMLBody + XHMTL.CreateCell()
                    HTMLBody = HTMLBody + Tools.LastComment(issue)
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # ********************************
                    # Insert Priority
                    # ********************************
                    HTMLBody = HTMLBody + XHMTL.CreateCell()
                    HTMLBody = HTMLBody + str(issue.fields.priority)
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # End Row
                    HTMLBody = HTMLBody + XHMTL.CloseRow()

                    print('Inserted Issue ' + issue.key)

                # CloseTable
                HTMLBody = HTMLBody + XHMTL.CloseTable()

                # And 2 lines return
                for space in range(2):
                    HTMLBody = HTMLBody + "<p><br /></p>"

            # Write Page in Confluence
            Tools.WriteConfluence('ASSIGNEE-PAGE', self.ArrayAssignee[i], HTMLBody)
        return;

    # ********************************
    # Pending Features Report
    # ********************************
    def PendingFeatures(self):

        # ********************************
        # Create Table
        # ********************************

        HTMLBody = XHMTL.CreateTable()

        # ********************************
        # Create Columns
        # ********************************

        ColName = ['Team', 'Features', 'Stories', 'Dependencies', 'Assignee', 'Sprint', 'Story Points',
                   'Comments', 'Priority']

        # New Row
        HTMLBody = HTMLBody + XHMTL.CreateRow()

        for Col in range(len(ColName)):
            HTMLBody = HTMLBody + XHMTL.CreateColumn()
            HTMLBody = HTMLBody + XHMTL.InsertColName(ColName[Col])
            HTMLBody = HTMLBody + XHMTL.CloseColumn()

        HTMLBody = HTMLBody + XHMTL.CloseRow()

        # End Columns

        # ********************************
        # Init Cont
        # ********************************

        # Initiate New Line Cont
        Newline = 0
        TotFeat = 0
        ContStory = 0
        TotFeatTeam = 0
        TotStoryTeam = 0
        TotStories = 0

        # ********************************
        # Add Rows
        # ********************************

        # Create a New Row for each Team
        for i in range(len(self.ArrayTeam)):

            # Query JIRA
            if self.ArrayTeam[i] != "None":
                Query = config['INI'].get('FeatureJQL') + ' AND resolution = Unresolved' + ' AND "Assigned Team"="' + \
                        self.ArrayTeam[i] + '"'
            else:
                Query = config['INI'].get('FeatureJQL') + ' AND resolution = Unresolved' +  ' AND "Assigned Team" is EMPTY'
            # Add Order By
            Query = Query + " " + config['INI'].get('FeatORDER')

            #Init Feature Cont
            ContFeat = 0
            TotFeatTeam = 0
            TotStoryTeam = 0

            print(Query)

            # Execute Feature Query in JIRA
            Issues = jira.search_issues(Query, startAt=0, maxResults=False)

            # Skip if the team do not have Features
            if len(Issues) > 0:
                # New Row
                HTMLBody = HTMLBody + XHMTL.CreateRow()

                # ********************************
                # Insert Team
                # ********************************
                HTMLBody = HTMLBody + XHMTL.CreateMergedCell("Flapuda")
                HTMLBody = HTMLBody + Tools.CleanTxt(self.ArrayTeam[i])
                HTMLBody = HTMLBody + XHMTL.CloseCell()

                for issue in Issues:
                    if ContFeat > 0:
                        # New Row
                        HTMLBody = HTMLBody + XHMTL.CreateRow()

                    #Get FixVersion
                    FixVersion = Tools.GetFixVersion(issue)

                    # Stories per Feature
                    StoryQuery = config['INI'].get('StoryJQL') + " AND resolution = Unresolved AND 'Epic Link'=" + issue.key
                    StoryQuery = StoryQuery + " " + config['INI'].get('StoryORDER')

                    # Calculate Feature Story Points
                    storySP = Tools.SUMPoints(issue.key)

                    # Execute Story Query in JIRA
                    Stories = jira.search_issues(StoryQuery, startAt=0, maxResults=False)

                    # ********************************
                    # Insert Feature
                    # ********************************
                    HTMLBody = HTMLBody + XHMTL.CreateMergedCell("Featuda")
                    HTMLBody = HTMLBody + XHMTL.InsertIssue(issue.key)
                    HTMLBody = HTMLBody + XHMTL.CloseCell()

                    # Add new Feature line
                    ContFeat = ContFeat + 1
                    TotFeat = TotFeat + 1
                    TotFeatTeam = TotFeatTeam + 1

                    # ********************************
                    # Insert Stories
                    # ********************************
                    # Feature with no Stories
                    if len(Stories) == 0:
                        for r in range(7):
                            HTMLBody = HTMLBody + XHMTL.CreateCell()
                            HTMLBody = HTMLBody + " "
                            HTMLBody = HTMLBody + XHMTL.CloseCell()

                    for issues in Stories:

                        if ContStory > 0:
                            # New Row
                            HTMLBody = HTMLBody + XHMTL.CreateRow()

                        # Create Cell
                        HTMLBody = HTMLBody + XHMTL.CreateCell()
                        HTMLBody = HTMLBody + XHMTL.InsertIssue(issues.key)
                        # Close Story Cell
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        # Add one new story to the cont
                        TotStories = TotStories + 1

                        # ********************************
                        # Insert Dependencies
                        # ********************************

                        HTMLBody = Tools.InsertDependencies(HTMLBody, issues)

                        # ********************************
                        # Insert Assignee
                        # ********************************
                        HTMLBody = HTMLBody + XHMTL.CreateCenterCell()
                        try:
                            HTMLBody = HTMLBody + str(issues.fields.assignee.displayName)
                        except:
                            HTMLBody = HTMLBody + "<p>N/A</p>"
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        # ********************************
                        # Insert Sprint
                        # ********************************
                        HTMLBody = HTMLBody + XHMTL.CreateCenterCell()
                        HTMLBody = HTMLBody + "<p>" + Tools.FindSprint(issues) + "</p>"
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        # ********************************
                        # Insert Story Points
                        # ********************************
                        HTMLBody = HTMLBody + XHMTL.CreateCenterCell()
                        HTMLBody = HTMLBody + str(issues.fields.customfield_10002)
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        # ********************************
                        # Insert Comments
                        # ********************************
                        HTMLBody = HTMLBody + XHMTL.CreateCell()
                        #HTMLBody = HTMLBody + Tools.LastComment(issue)
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        # ********************************
                        # Insert Priority
                        # ********************************
                        HTMLBody = HTMLBody + XHMTL.CreateCenterCell()
                        HTMLBody = HTMLBody + str(issues.fields.priority)
                        HTMLBody = HTMLBody + XHMTL.CloseCell()

                        #Add 1 Story Row
                        ContStory  = ContStory + 1

                        #End of Row
                        HTMLBody = HTMLBody + XHMTL.CloseRow()
                        Newline = Newline  + 1

                    #Init Cont Stories
                    if ContStory > 0:
                        HTMLBody = HTMLBody.replace("Featuda", str(ContStory))
                        TotStoryTeam = TotStoryTeam + ContStory
                        ContStory = 0
                    else:
                        #End of Row
                        HTMLBody = HTMLBody + XHMTL.CloseRow()
                        Newline = Newline + 1

                    #print('Inserted Feature ' + issue.key)

                if ContFeat == 0:
                    # End of Row
                    HTMLBody = HTMLBody + XHMTL.CloseRow()
                    Newline = Newline + 1
                else:
                    HTMLBody = HTMLBody.replace("Flapuda", str(Newline))
                    Newline = 0

                # Number of Features per team
                print("Total Features for " + str(self.ArrayTeam[i]) + ": " + str(TotFeatTeam))
                # Number of Stories per team
                print("Total Stories for " + str(self.ArrayTeam[i]) + ": " + str(TotStoryTeam))

        # CloseTable
        HTMLBody = HTMLBody + XHMTL.CloseTable()
        print("Total Features: " + str(TotFeat))
        print("Total Stories: " + str(TotStories))

        # Write Page in Confluence
        Tools.WriteConfluence('PENDING-PAGE',"VMH - ", HTMLBody)

        return;
