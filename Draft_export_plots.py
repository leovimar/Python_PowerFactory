import os
import powerfactory
app = powerfactory.GetApplication()

path  = 'D:\Projects\Gobel\Results_'
folder = 'Plots'


'''
Some text goes here
'''

'''
Some text goes here
'''

'''
Some text goes here
'''

'''
Some text goes here
'''


#folder_path = os.path.join(path, folder)  
folder_path = '%s\%s' %(path, folder)    

#os.mkdir(folder_path)

graph_board = app.GetGraphicsBoard()

page = 'my_file'
ext = 'pdf'

total = f'{page}.{ext}'

app.PrintPlain(total)

#def export_grid_pages(app, folder_path):
#  graph_board = app.GetGraphicsBoard()
#  grid_pages = graph_board.GetContents('*.SetDeskpage')
#  for grid_page in grid_pages:
#    grid_page.Show()
#    grid_page_name = grid_page.GetAttribute('loc_name')
#    file_name = '%s\%s' %(folder_path, grid_page_name)
##    file_name = os.path.join(folder_path, f'{grid_page.GetAttribute("loc_name")}')
#    graph_board.WriteWMF(file_name)
#
#def export_plot_pages(app, folder_path):
#  graph_board = app.GetGraphicsBoard()
#  plot_pages = graph_board.GetContents('*.GrpPage')
#  for plot_page in plot_pages:
#    plot_page.Show()
#    plot_page_name = plot_page.GetAttribute('loc_name')
#    file_name = '%s\%s' %(folder_path, plot_page_name)
#    graph_board.WriteWMF(file_name) 
     

'''
Extra information for testing Git
'''

#==============================================================================
# Other formats (emf, svg, pdf, png, jpg, ...)
#==============================================================================

def export_plot_pages(app, folder_path, extension):
  com_wr = app.GetFromStudyCase('ComWr')
  com_wr.iopt_rd = extension
  com_wr.iopt_savas = 0 # 0 = write to path, 1 = open save dialog
  plot_pages = graph_board.GetContents('*.GrpPage')
  for plot_page in plot_pages:
    plot_page.Show()
    plot_page_name = plot_page.GetAttribute('loc_name')
    file_name = f'{folder_path}\{plot_page_name}.{extension}' 
    graph_board.WriteWMF(file_name)
    com_wr.f = file_name
    com_wr.Execute()  

#iopt_rd = 'pdf'
#
## Get and set export command
#ComWr = app.GetFromStudyCase('ComWr')
#ComWr.SetAttribute('iopt_rd', iopt_rd)
#ComWr.SetAttribute('iopt_savas', 0) # 0 = write to path, 1 = open save dialog
#
## Exporting grid pages 
#for grid_page in graph_board.GetContents('*.SetDeskpage'):
#    grid_page.Show()
#    file_name = os.path.join(folder_path, f'{grid_page.GetAttribute("loc_name")}.{iopt_rd}')
#    ComWr.SetAttribute('f', file_name)
#    ComWr.Execute()
#
## Exporting plot pages
#for plot_page in graph_board.GetContents('*.GrpPage'):
#    plot_page.Show()
#    file_name = os.path.join(folder_path, f'{plot_page.GetAttribute("loc_name")}.{iopt_rd}')
#    ComWr.SetAttribute('f', file_name)
#    ComWr.Execute()

#export_grid_pages(app, folder_path)
#export_plot_pages(app, folder_path)

extension = 'pdf'

export_plot_pages(app, folder_path, extension)


