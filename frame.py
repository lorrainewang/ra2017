# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Dec 21 2016)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import wx.grid

###########################################################################
## Class Frame
###########################################################################

class Frame ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 575,700 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		
		sizer1 = wx.BoxSizer( wx.VERTICAL )
		
		self.notebook = wx.Notebook( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.notebook.SetFont( wx.Font( 10, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Sans" ) )
		
		self.panel1 = wx.Panel( self.notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		main = wx.BoxSizer( wx.VERTICAL )
		
		self.lbl_subjMapping = wx.StaticText( self.panel1, wx.ID_ANY, u"Subject Mapping", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.lbl_subjMapping.Wrap( -1 )
		self.lbl_subjMapping.SetFont( wx.Font( 9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Sans" ) )
		
		main.Add( self.lbl_subjMapping, 0, wx.ALL, 5 )
		
		self.fp_smpicker = wx.FilePickerCtrl( self.panel1, wx.ID_ANY, wx.EmptyString, u"Select Subject Mapping File", u"*.*", wx.DefaultPosition, wx.Size( 500,-1 ), wx.FLP_DEFAULT_STYLE )
		main.Add( self.fp_smpicker, 0, wx.ALL, 5 )
		
		self.m_staticline1 = wx.StaticLine( self.panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL )
		main.Add( self.m_staticline1, 0, wx.EXPAND |wx.ALL, 5 )
		
		boxSizer2 = wx.BoxSizer( wx.VERTICAL )
		
		grid_batch_subject = wx.GridSizer( 0, 2, 0, 0 )
		
		boxSizer_batch = wx.BoxSizer( wx.VERTICAL )
		
		self.lbl_batch = wx.StaticText( self.panel1, wx.ID_ANY, u"Batch", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.lbl_batch.Wrap( -1 )
		self.lbl_batch.SetFont( wx.Font( 9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Sans" ) )
		
		boxSizer_batch.Add( self.lbl_batch, 0, wx.ALL, 5 )
		
		listbox_batchChoices = []
		self.listbox_batch = wx.ListBox( self.panel1, wx.ID_ANY, wx.DefaultPosition, wx.Size( 200,150 ), listbox_batchChoices, 0 )
		boxSizer_batch.Add( self.listbox_batch, 0, wx.ALL, 5 )
		
		
		grid_batch_subject.Add( boxSizer_batch, 1, wx.EXPAND, 5 )
		
		boxSizer_subject = wx.BoxSizer( wx.VERTICAL )
		
		self.lbl_subj = wx.StaticText( self.panel1, wx.ID_ANY, u"Subject", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.lbl_subj.Wrap( -1 )
		self.lbl_subj.SetFont( wx.Font( 9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Sans" ) )
		
		boxSizer_subject.Add( self.lbl_subj, 0, wx.ALL, 5 )
		
		listbox_subjectChoices = []
		self.listbox_subject = wx.ListBox( self.panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, listbox_subjectChoices, 0 )
		self.listbox_subject.SetMinSize( wx.Size( 200,150 ) )
		
		boxSizer_subject.Add( self.listbox_subject, 0, wx.ALL, 5 )
		
		
		grid_batch_subject.Add( boxSizer_subject, 1, wx.EXPAND, 5 )
		
		
		boxSizer2.Add( grid_batch_subject, 1, wx.EXPAND, 5 )
		
		fgSizer1 = wx.FlexGridSizer( 0, 2, 0, 0 )
		fgSizer1.SetFlexibleDirection( wx.HORIZONTAL )
		fgSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		boxSizer_ItemDetails = wx.BoxSizer( wx.VERTICAL )
		
		self.lbl_itemdetails = wx.StaticText( self.panel1, wx.ID_ANY, u"Item Details", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.lbl_itemdetails.Wrap( -1 )
		self.lbl_itemdetails.SetFont( wx.Font( 9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Sans" ) )
		
		boxSizer_ItemDetails.Add( self.lbl_itemdetails, 0, wx.ALL, 5 )
		
		self.grid_non_mcq_questions = wx.grid.Grid( self.panel1, wx.ID_ANY, wx.DefaultPosition, wx.Size( 400,350 ), 0 )
		
		# Grid
		self.grid_non_mcq_questions.CreateGrid( 50, 2 )
		self.grid_non_mcq_questions.EnableEditing( True )
		self.grid_non_mcq_questions.EnableGridLines( True )
		self.grid_non_mcq_questions.EnableDragGridSize( False )
		self.grid_non_mcq_questions.SetMargins( 0, 0 )
		
		# Columns
		self.grid_non_mcq_questions.SetColSize( 0, 150 )
		self.grid_non_mcq_questions.SetColSize( 1, 139 )
		self.grid_non_mcq_questions.EnableDragColMove( False )
		self.grid_non_mcq_questions.EnableDragColSize( True )
		self.grid_non_mcq_questions.SetColLabelSize( 30 )
		self.grid_non_mcq_questions.SetColLabelValue( 0, u"Item Description" )
		self.grid_non_mcq_questions.SetColLabelValue( 1, u"Item Max. Marks" )
		self.grid_non_mcq_questions.SetColLabelAlignment( wx.ALIGN_CENTRE, wx.ALIGN_CENTRE )
		
		# Rows
		self.grid_non_mcq_questions.EnableDragRowSize( True )
		self.grid_non_mcq_questions.SetRowLabelSize( 80 )
		self.grid_non_mcq_questions.SetRowLabelAlignment( wx.ALIGN_CENTRE, wx.ALIGN_CENTRE )
		
		# Label Appearance
		self.grid_non_mcq_questions.SetLabelTextColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOWFRAME ) )
		
		# Cell Defaults
		self.grid_non_mcq_questions.SetDefaultCellAlignment( wx.ALIGN_LEFT, wx.ALIGN_TOP )
		boxSizer_ItemDetails.Add( self.grid_non_mcq_questions, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( boxSizer_ItemDetails, 1, wx.EXPAND, 5 )
		
		boxSizer_Buttons = wx.BoxSizer( wx.VERTICAL )
		
		
		#boxSizer_Buttons.AddSpacer(  0, 0, 1, wx.EXPAND, 5 )
		#boxSizer_Buttons.AddSpacer( 0, 1, wx.EXPAND, 5 )
		
		self.btn_genTemplate = wx.Button( self.panel1, wx.ID_ANY, u"Generate\nTemplate", wx.DefaultPosition, wx.DefaultSize, 0 )
		boxSizer_Buttons.Add( self.btn_genTemplate, 0, wx.ALL, 5 )
		
		
		#boxSizer_Buttons.AddSpacer(  0, 0, 1, wx.EXPAND, 5 )
		#boxSizer_Buttons.AddSpacer( 0, 1, wx.EXPAND, 5 )
		
		self.btn_clearItems = wx.Button( self.panel1, wx.ID_ANY, u"Clear", wx.DefaultPosition, wx.DefaultSize, 0 )
		boxSizer_Buttons.Add( self.btn_clearItems, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( boxSizer_Buttons, 1, wx.EXPAND, 5 )
		
		
		boxSizer2.Add( fgSizer1, 1, wx.EXPAND, 5 )
		
		
		main.Add( boxSizer2, 1, wx.EXPAND, 5 )
		
		
		self.panel1.SetSizer( main )
		self.panel1.Layout()
		main.Fit( self.panel1 )
		self.notebook.AddPage( self.panel1, u"Generate Student List", True )
		self.panel2 = wx.Panel( self.notebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		gbSizer1 = wx.GridBagSizer( 0, 0 )
		gbSizer1.SetFlexibleDirection( wx.BOTH )
		gbSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_ALL )
		
		bSizer11 = wx.BoxSizer( wx.VERTICAL )
		
##		self.m_staticText11 = wx.StaticText( self.panel2, wx.ID_ANY, u"Import MCQ and Non-MCQ Items", wx.DefaultPosition, wx.DefaultSize, 0 )
##		self.m_staticText11.Wrap( -1 )
##		self.m_staticText11.SetFont( wx.Font( 10, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False, "Sans" ) )
		
		
		
		sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self.panel2, wx.ID_ANY, u' ►► Import Assessment Items ◄◄'), wx.VERTICAL )
		
		bSizer91 = wx.BoxSizer( wx.VERTICAL )
		#bSizer91.Add( self.m_staticText11, 0, wx.ALL, 5 )
		
		bSizer18 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.checkbox_include_mcq = wx.CheckBox( sbSizer3.GetStaticBox(), wx.ID_ANY, u"Include MCQ Responses", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.checkbox_include_mcq.SetValue(True) 
		bSizer18.Add( self.checkbox_include_mcq, 0, wx.ALL, 5 )
		
		
		
		
		bSizer21 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.text_mcq_responses_dir = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"MCQ Responses Directory", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.text_mcq_responses_dir.Wrap( -1 )
		bSizer21.Add( self.text_mcq_responses_dir, 0, wx.ALL, 5 )
		
		self.dirpicker_mcq_responses = wx.DirPickerCtrl( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select the MCQ Responses Folder", wx.DefaultPosition, wx.DefaultSize, style = wx.DIRP_DEFAULT_STYLE )
		bSizer21.Add( self.dirpicker_mcq_responses, 0, wx.ALL, 5 )
		
		
		bSizer211 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.text_mcq_answers = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"MCQ Answers", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.text_mcq_answers.Wrap( -1 )
		bSizer211.Add( self.text_mcq_answers, 0, wx.ALL, 5 )
		
		self.filepicker_mcq_answers = wx.FilePickerCtrl( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select the Answer Key File", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		bSizer211.Add( self.filepicker_mcq_answers, 0, wx.ALL, 5 )
		
		
			
		
		sbSizer3.Add( bSizer91, 1, wx.EXPAND, 5 )
		
		
		bSizer11.Add( sbSizer3, 1, wx.EXPAND, 5 )
		
		sbSizer31 = wx.StaticBoxSizer( wx.StaticBox( self.panel2, wx.ID_ANY, u"Non MCQ Question" ), wx.VERTICAL )
		
		#bSizer911 = wx.BoxSizer( wx.VERTICAL )
		
		#bSizer181 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.checkbox_include_non_mcq = wx.CheckBox( sbSizer3.GetStaticBox(), wx.ID_ANY, u"Include Non-MCQ Responses in Student List file", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.checkbox_include_non_mcq.SetValue(True) 
		
		
		
		#bSizer911.Add( bSizer181, 1, wx.EXPAND, 5 )
		
		bSizer2111 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.text_non_mcq_questions = wx.StaticText( sbSizer3.GetStaticBox(), wx.ID_ANY, u"Student List", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.text_non_mcq_questions.Wrap( -1 )
		bSizer2111.Add( self.text_non_mcq_questions, 0, wx.ALL, 5 )
		
		self.filepicker_non_mcq_questions = wx.FilePickerCtrl( sbSizer3.GetStaticBox(), wx.ID_ANY, wx.EmptyString, u"Select file containing Student List", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		bSizer2111.Add( self.filepicker_non_mcq_questions, 0, wx.ALL, 5 )
		
		
		bSizer91.Add( bSizer2111, 1, wx.EXPAND, 5 )
		bSizer91.Add( bSizer18, 1, wx.EXPAND, 5 )
		bSizer91.Add( bSizer21, 1, wx.EXPAND, 5 )
		bSizer91.Add( bSizer211, 1, wx.EXPAND, 5 )
		bSizer91.Add( self.checkbox_include_non_mcq, 0, wx.ALL, 5 )

		
		#sbSizer3.Add( bSizer911, 1, wx.EXPAND, 5 )
		
		
		#bSizer11.Add( sbSizer31, 1, wx.EXPAND, 5 )
		
		bSizer16 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.m_staticText10 = wx.StaticText( self.panel2, wx.ID_ANY, u"Highlight all MCQ response percentages below ", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText10.Wrap( -1 )
		bSizer16.Add( self.m_staticText10, 0, wx.ALL, 5 )
		
		self.input_mcq_reponse_percentage_highlight_threshold = wx.TextCtrl( self.panel2, wx.ID_ANY, u"5", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer16.Add( self.input_mcq_reponse_percentage_highlight_threshold, 0, wx.ALL, 5 )
		
		self.m_staticText111 = wx.StaticText( self.panel2, wx.ID_ANY, u"% red", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText111.Wrap( -1 )
		bSizer16.Add( self.m_staticText111, 0, wx.ALL, 5 )
		
		
		bSizer11.Add( bSizer16, 0, wx.EXPAND, 10 )
		
		self.btn_mergeAndAnalyse = wx.Button( self.panel2, wx.ID_ANY, u"Merge and Analyse", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer11.Add( self.btn_mergeAndAnalyse, 0, wx.ALL, 5 )
		
		
		gbSizer1.Add( bSizer11, wx.GBPosition( 0, 0 ), wx.GBSpan( 1, 1 ), wx.EXPAND, 5 )
		
		
		self.panel2.SetSizer( gbSizer1 )
		self.panel2.Layout()
		gbSizer1.Fit( self.panel2 )
		self.notebook.AddPage( self.panel2, u"Item Analysis", False )
		
		sizer1.Add( self.notebook, 1, wx.EXPAND |wx.ALL, 5 )
		
		
		self.SetSizer( sizer1 )
		self.Layout()
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.fp_smpicker.Bind( wx.EVT_FILEPICKER_CHANGED, self.subject_mapping_changed )
		self.listbox_batch.Bind( wx.EVT_LISTBOX_DCLICK, self.batch_changed )
		self.listbox_subject.Bind( wx.EVT_LISTBOX_DCLICK, self.populate_student_details )
		self.btn_genTemplate.Bind( wx.EVT_BUTTON, self.generate_template )
		self.btn_clearItems.Bind( wx.EVT_BUTTON, self.clear_non_mcq_questions )
		self.checkbox_include_mcq.Bind( wx.EVT_CHECKBOX, self.include_mcq_toggled )
		self.dirpicker_mcq_responses.Bind( wx.EVT_DIRPICKER_CHANGED, self.mcq_dir_changed )
		self.filepicker_mcq_answers.Bind( wx.EVT_FILEPICKER_CHANGED, self.mcq_answer_template_changed )
		self.checkbox_include_non_mcq.Bind( wx.EVT_CHECKBOX, self.include_non_mcq_toggled )
		self.filepicker_non_mcq_questions.Bind( wx.EVT_FILEPICKER_CHANGED, self.non_mcq_questions_changed )
		self.btn_mergeAndAnalyse.Bind( wx.EVT_BUTTON, self.merge_and_analyse )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def subject_mapping_changed( self, event ):
		event.Skip()
	
	def batch_changed( self, event ):
		event.Skip()
	
	def populate_student_details( self, event ):
		event.Skip()
	
	def generate_template( self, event ):
		event.Skip()
	
	def clear_non_mcq_questions( self, event ):
		event.Skip()
	
	def include_mcq_toggled( self, event ):
		event.Skip()
	
	def mcq_dir_changed( self, event ):
		event.Skip()
	
	def mcq_answer_template_changed( self, event ):
		event.Skip()
	
	def include_non_mcq_toggled( self, event ):
		event.Skip()
	
	def non_mcq_questions_changed( self, event ):
		event.Skip()
	
	def merge_and_analyse( self, event ):
		event.Skip()
