def GetContext(text, context, tags):
    '''Finds those words in the given text which are between the tag symbols
    and adds them to the context dictionary.
    Those words represent the parts of the document that need to be filled with information.'''
    tag, gat = tags[:len(tags)//2], tags[len(tags)//2:]
    if text != '':
        i1, i2 = 0, 0
        for i in range(0, len(text)-len(tag)):
            if text[i:i+len(tag)] == tag:
                i1 = i + len(tag)
                for j in range(i+len(tag), len(text)-len(tag)+1):
                    if text[j:j+len(tag)] == gat:
                        i2 = j - len(tag) + 1
        if i1 != 0 and i2 != 0:
            c = text[i1:i2+1].strip()
            if c not in context.keys():
                context[c] = ''

import tkinter as tk
import tkinter.filedialog
from docxtpl import DocxTemplate
class TemplateFiller(tk.Frame):
    def __init__(self, parent):
        '''Creates GUI.'''
        tk.Frame.__init__(self, parent) #?
        self.window = parent #?
        self.context, self.context_vars = {}, {} #stores the parts that need to be filled and the tk.Entry variables
        #title
        tk.Label(self.window, text='Template filler', font=('Helvetica 15 bold')).grid(row=0, column=0)
        #load template
        tk.Button(self.window, text='Load template', command=self.load_template).grid(row=1, column=0)
        #default context (doesn't work yet)
        choices = {''}
        self.defc = tk.StringVar()
        self.defc.set('Choose default context')
        self.defc_menu = tk.OptionMenu(self.window, self.defc, *choices)
        self.defc_menu.grid(row=1, column=1)
        #get context from the loaded template
        tk.Button(self.window, text='Get context', command=self.get_context).grid(row=2, column=0)
        #render context to a new document
        tk.Button(self.window, text='Render context', command=self.render_context).grid(row=2, column=1)
        
    def load_template(self):
        self.template_filename = tk.filedialog.askopenfilename(title='Load template file')
        x = self.template_filename.split('/')
        self.filename = x[len(x)-1] #filename is the part of path after the last / symbol
        self.template = DocxTemplate(self.template_filename) #creates a DocxTemplate object

    def get_context(self):
        '''Goes over the template to find the context
        (the words that are between the tags)'''
        #paragraph parsing
        for para in self.template.paragraphs: #for every paragraph in the template
            txt = para.text
            GetContext(txt, self.context, tags='{{}}') #get the context from the text
        #table parsing
        for table in self.template.tables: #for every table in the template
            for row in table.rows: #for every row in the table
                for cell in row.cells: #for every cell in the row
                    for para in cell.paragraphs: #for every paragraph in the cell
                        txt = para.text
                        GetContext(txt, self.context, tags='{{}}') #get the context from the text
        #image parsing
        #to be made
        self.show_context()

    def show_context(self):
        '''Shows a layout of the template's context.
        Based on this layout the document can be filled with the required information.'''
        try:
            self.context_area.destroy()
        except:
            pass
        self.context_area = tk.Frame(self.window)
        self.context_area.grid(row=3, column=0, columnspan=2)
        tk.Label(self.context_area, text='save filename').grid(row=0, column=0) #the name ot the new document
        self.save_filename = tk.StringVar()
        self.save_filename.set(self.filename)
        tk.Entry(self.context_area, textvariable=self.save_filename, width=30).grid(row=0, column=1) #an entry for the filename
        r = 1
        for c in self.context.keys(): #for every word in the context
            tk.Label(self.context_area, text=c).grid(row=r, column=0) #create a label
            c_var = tk.StringVar()
            self.context_vars[c] = c_var
            tk.Entry(self.context_area, textvariable=self.context_vars[c], width=30).grid(row=r, column=1) #create an entry for information
            r += 1

    def render_context(self):
        '''Creates a new document with the given name and informations based on the loaded template.'''
        for c in self.context_vars.keys(): #for every word in the context
            self.context[c] = self.context_vars[c].get() #read the information from the given Entry
        self.template.render(self.context) #render the context into the template
        self.template.save(self.save_filename.get()) #saves the template
        
if __name__ == '__main__':
    root = tk.Tk()
    TF = TemplateFiller(root)
    TF.grid(row=0, column=0)
    root.mainloop()
