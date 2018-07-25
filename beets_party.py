import pandas as pd
import os
import matplotlib.pyplot as plt
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

os.chdir('//Dc1nas/projects-ab/AS/Data-OPS/Hydra Data/Intern Hydra Project 2018')

df = pd.read_excel('TrueCalBrazeData1.xlsx', header=0, index=0)
df_limits = pd.read_excel('TrueCalBrazeData1.xlsx', sheet_name='Spec Limits', header=0, index_col=0)

#Excluding REF, PROFILE, and Symmetry dimensions
df = df.loc[:,[x for x in df.columns if x.count('REF') == 0 and x.count('PROFILE')==0 and x.count('Symmetry')==0]]
df_limits = df_limits.loc[[x for x in df_limits.index if x.count('REF') == 0 and x.count('PROFILE')==0 and x.count('Symmetry')==0],:]
#Excluding all but raw data
df1 = df.iloc[:-1,3:]

#Extracting specification limits from manually-made spec sheet
nom_vals = df_limits.iloc[:,0]
tol_vals = df_limits.iloc[:,1]

#df_long = df.iloc[:-1,2:]
#df_long = df_long.reset_index()
#df_long = pd.melt(df_long, id_vars=['index','Lot Number'])

# =============================================================================
# Plotting for dimensions that have Cpk values < 1
# =============================================================================
#Calculating Cpk for all dimensions
df_limits['Cpl'] = [(df1.iloc[:,i].mean() - (nom_vals[i] - tol_vals[i]))/(3*df1.iloc[:,i].std()) for i in range(len(df_limits))]
df_limits['Cpu'] = [((nom_vals[i] + tol_vals[i]) - df1.iloc[:,i].mean())/(3*df1.iloc[:,i].std()) for i in range(len(df_limits))]
df_limits['Cpk'] = df_limits[['Cpl','Cpu']].min(axis=1)

print("     Plots for dimensions that have Cpk values < 1:")

for i in range(len(df_limits)):
    if df_limits['Cpk'].iloc[i] < 3:
        
        #Histograms
#        plt.subplot(211)
        plt_names = 'histogram - ' + str(i) + '.png' 
        plt.hist(df1.iloc[:,i], bins=25)
        plt.axvline(x=nom_vals.iloc[i], color='k', label='Nominal Value: {}'.format(nom_vals[i]), linestyle='--')
        plt.axvline(x=nom_vals[i]+tol_vals[i], color='k', label='+/- {} Tolerance'.format(tol_vals[i]))
        plt.axvline(x=nom_vals[i]-tol_vals[i], color='k')
        plt.title('Histogram: {} - Incapable'.format(df1.columns[i]))
        plt.ylabel('Frequency')
        plt.xlabel(df1.columns[i])
        plt.legend()
        plt.savefig('histogram - ' + str(i) + '.png')
        plt.show()
        
        #Control Charts
#        plt.subplot(212)
        plt.plot(df1.iloc[:,i])
        plt.axhline(y=nom_vals[i]+tol_vals[i], color='r', label='USL/LSL')
        plt.axhline(y=nom_vals[i]-tol_vals[i], color='r')
        plt.axhline(y=nom_vals[i], color='k', label='Process Mean', linestyle='--')
        plt.axhline(y=df1.iloc[:,i].std()*3+nom_vals[i], color='c', label='UCL/LCL - 3 sigma', linestyle=':')
        plt.axhline(y=nom_vals[i]-df1.iloc[:,i].std()*3, color='c', linestyle=':')
        plt.legend()
        plt.ylabel(df1.columns[i])
        plt.xlabel('Count')
        plt.title('Time Series Control Chart - {}'.format(df1.columns[i]))
       
        for j in range(len(df1)):
            #Marking points more extreme than 3-sigma
            if abs(nom_vals[i]-df1.iloc[j,i]) > 3*df1.iloc[:,i].std():
                plt.plot(j,df1.iloc[j,i], 'co')
            #Marking points more extreme than spec limits
            if abs(nom_vals[i]-df1.iloc[j,i]) > tol_vals[i]:
                plt.plot(j,df1.iloc[j,i], 'ro')
        plt.savefig('control chart - ' + str(i) + '.png')
        plt.show()

# =============================================================================
# Dimension Summary Information Data Frame                
# =============================================================================
dimsum_cols = ['Spec Limits', 'Process Capability', 'OOC Counts']
dimsum_index = df_limits.index
df_dimsum = pd.DataFrame(index=dimsum_index, columns=dimsum_cols)

df_dimsum['Spec Limits'] = ['{} +/- {}'.format(nom_vals[i], tol_vals[i]) for i in range(len(df_limits))]

#Proces Capability Categorization
cp_category = []
for i in df_limits['Cpk']:
    if i < 1:
        cp_category += ['Incapable']
    elif i < 2:
        cp_category += ['Capable']
    else:
        cp_category += ['Six-Sigma Capable']
df_dimsum['Process Capability'] = cp_category

#Counts at which data is out of spec
ooc_points = []
for i in range(len(df1.columns)):
    ooc_points_indv = []
    for j in enumerate(df1.iloc[:,i]):
        #If exactly on spec limit, considered out-of-control
        if abs(nom_vals[i]-df1.iloc[j[0],i]) >= tol_vals[i]:
            ooc_points_indv += [j[0]+1]
    if len(ooc_points_indv) == 0:
        ooc_points_indv += [None]
    ooc_points += [ooc_points_indv]
df_dimsum['OOC Counts'] = [", ".join(repr(x) for x in ooc_points[i]) for i in range(len(ooc_points))]

# =============================================================================
# Individual Part Summary Information Data Frame
# =============================================================================
partsum_cols = ['OOC Dimensions']
partsum_index = df['COUNT'][:-1]
df_partsum = pd.DataFrame(index=partsum_index, columns=partsum_cols)
df_partsum.index.names = ['Part Count #']

#Dimensions out of spec for each part                          
ooc_dims = []
for i in range(len(df1)):
    ooc_dims_indv = []
    for j in range(len(df1.columns)):
        if abs(nom_vals[j]-df1.iloc[i,j]) >= tol_vals[j]:
            ooc_dims_indv += [df1.columns[j]]
    if len(ooc_dims_indv) == 0:
        ooc_dims_indv += [None]
    ooc_dims += [ooc_dims_indv]
df_partsum['OOC Dimensions'] = [', '.join(repr(x) for x in ooc_dims[i]) for i in range(len(ooc_dims))]
#Filtering out parts not out-of-control
df_partsum = df_partsum[df_partsum['OOC Dimensions'] != 'None']

# =============================================================================
# To PDF            
# =============================================================================

env = Environment(loader = FileSystemLoader('.'))
template = env.get_template("beets_html.html")
template_vars = {"title": 'Supplier Dimension Data Summary',
                 'df_dimsum': df_dimsum.to_html(),
                 'df_partsum': df_partsum.to_html(),
                 'histogram': plt_names}



html_out = template.render(template_vars)

HTML(string=html_out).write_pdf("deep_purple_example_output.pdf")