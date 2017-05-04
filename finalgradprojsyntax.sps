* Encoding: UTF-8.

GET
  FILE='/Users/daniellecrumley/Desktop/Graduate Project/SPSS data for CSCI E-7 proj.sav'.
DATASET NAME DataSet1 WINDOW=FRONT.

DATASET ACTIVATE DataSet1.
REGRESSION
  /MISSING LISTWISE
  /STATISTICS COEFF OUTS CI(95) R ANOVA
  /CRITERIA=PIN(.05) POUT(.10)
  /NOORIGIN 
  /DEPENDENT MentalHealthScore
  /METHOD=ENTER MonthsinVN ProblemOnDischarge_Health ProblemOnDischarge_Money 
    ProblemOnDischarge_Job BiggestEvent VNDummy1 VNDummy2 EducDummy1 EducDummy2 Age 
    exposurekilling_binary.



begin program python3.
import SpssClient
SpssClient.StartClient()
oDoc = SpssClient.GetDesignatedOutputDoc()  #Access active output window
oItems = oDoc.GetOutputItems() #Look up all items in output window

for index in range(oItems.Size()): #Loop through indices of output items
    oItem = oItems.GetItemAt(index) #Access each output item
    if oItem.GetType() == SpssClient.OutputItemType.PIVOT: #Continue only if output item is pivot table
        pTable = oItem.GetSpecificType() #Access pivot table
        datacells = pTable.DataCellArray()
        PivotMgr = pTable.PivotManager()
        ColLabels = pTable.ColumnLabelArray()

#pull information from Model Summary Table and ANOVA Table (for footnote) and from Coefficients Table
        if pTable.GetTitleText() == 'Model Summary':
            modelsuminfo = {} #keys will be the names of the stats we need
			    #'R Square' and 'Adjusted R Square
			    #values will be the corresponding value of each statistic
            for i in range(1,ColLabels.GetNumRows()):#iterating through all rows and columns of the Column Labels
                for j in range(ColLabels.GetNumColumns()):
                    label = ColLabels.GetValueAt(i,j)
                    if label in ['R Square', 'Adjusted R Square']:#the stats we are interested in
                        lastrowindex = datacells.GetNumRows() - 1 #in this case there's only one row of data
				#but in diff regression models there can be multiple rows, and last row
				#contains the final values we want
                        valuetoadd= float(datacells.GetUnformattedValueAt (lastrowindex, j))
				#the value of R Square or Adjusted R Square
                        valuetoadd = "{0:.2f}".format(valuetoadd)
                        modelsuminfo[label]=valuetoadd
        elif pTable.GetTitleText() == 'ANOVA':
		#now we gather the values from df, F, and Sig columns of ANOVA table
            anovavals = {}
            rownames = ['Regression', 'Residual']
            for i in range(1,ColLabels.GetNumRows()):
                for j in range(ColLabels.GetNumColumns()):
                    label = ColLabels.GetValueAt(i,j)
                    if label in ['df', 'F', 'Sig.']:
                        for k in range (PivotMgr.GetNumRowDimensions()):
                            RowDim = PivotMgr.GetRowDimension(k)
                            for m in range (RowDim.GetNumCategories()):
                                rowname = RowDim.GetCategoryValueAt (m)
                                if rowname in rownames:
                                    if datacells.GetUnformattedValueAt (m, j) == '':
                                        valuetoadd = ''
                                    else:
                                        valuetoadd= float(datacells.GetUnformattedValueAt (m, j))
                                        valuetoadd = "{0:.2f}".format(valuetoadd)
                                    if rowname not in anovavals:
                                        anovavals[rowname] = [[label, valuetoadd]]
						#label is the name of the statistic (eg., 'df')
						#valuetoadd is corresponding value
                                    else:
                                        anovavals[rowname].append ([label, valuetoadd])
        
        elif pTable.GetTitleText() == 'Coefficients': #gathering values from Coefficients Table
            stat_labels =  ['B', 'Std. Error', 'Beta', 'Sig.', 'Lower Bound', 'Upper Bound']
            coeffdict = {}
            lastrow = ColLabels.GetNumRows() - 1
            for j in range(ColLabels.GetNumColumns()):
                label = ColLabels.GetValueAt(lastrow,j)
                if label in stat_labels:
                    print ("this is the label we are working with:", label)
                    for i in range (PivotMgr.GetNumRowDimensions()):
                        RowDim = PivotMgr.GetRowDimension(i)
                        if RowDim.GetDimensionName() == 'Variables':
                            varnames = []
                            for k in range (RowDim.GetNumCategories()):
                                variable_name = RowDim.GetCategoryValueAt (k)
                                varnames.append (variable_name) #adds name of each variable in the model to varnames list
                                if datacells.GetUnformattedValueAt (k, j) == '':
                                    valuetoadd = ''
                                else:
                                    valuetoadd = float(datacells.GetUnformattedValueAt (k, j))
                                    valuetoadd = "{0:.2f}".format(valuetoadd)
                                if variable_name not in coeffdict:
                                    coeffdict[variable_name] = [[label, valuetoadd]]
                                else:
                                    coeffdict[variable_name]. append ([label, valuetoadd])

## editing coeff dict so that the value corresponding to B is actually B [lower bound of CI, upper bound of CI]
for key, value in coeffdict.items():
    oldBval = value[0][1]
    LowerBoundCoeff = value [4][1]
    UpperBoundCoeff = value [5][1]
    value [0][1] = oldBval + '\n[{}, {}]'.format (LowerBoundCoeff, UpperBoundCoeff)
    value.remove (value[5]) #removing Lower Bound and Upper Bound lists 
    value.remove (value[4])

stat_labels = ['B [95% CI]', 'SE B', 'Beta', 'p']
    #renaming the statistics labels in order

# footnote for our table - uses values from Model Summary and ANOVA tables
footnote = "Note: R Square = {}, adjusted R Square = {}, overall F (df {}, {}) = {}, p = {}.".format(\
modelsuminfo['R Square'], modelsuminfo['Adjusted R Square'], anovavals['Regression'][0][1], anovavals ['Residual'][0][1], \
anovavals ['Regression'][1][1], anovavals ['Regression'][2][1])



##finally, create new pivot table with the values we've pulled above
import spss
mytablecells = []
for var in varnames: # because the variable names are sorted as we want them in this list
    for i in range (len(coeffdict[var])):
        statdata = coeffdict[var][i][1]
        mytablecells.append (statdata)

spss.StartProcedure("myTABLE")
mytitle = 'Linear Model of Predictors of the Mental Health Score. 95% CI and standard errors based on parametric model.'
table = spss.BasePivotTable (mytitle, "OMS table subtype")
table.SimplePivotTable(rowdim = "Variables in Model",
                       rowlabels = varnames,
                       coldim = "Statistics",
                       collabels = stat_labels,
                       cells = mytablecells)
table.TitleFootnotes (footnote) # adds footnote to the title of the table
spss.EndProcedure() #then the table will print to the output
end program.

