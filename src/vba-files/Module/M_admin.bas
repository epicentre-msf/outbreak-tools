Attribute VB_Name = "M_admin"
Option Explicit

Sub exporterCode()

ThisWorkbook.VBProject.VBComponents("F_GEO").Export "D:\OneDrive - MSF\Documents\Source\F_Geo.frm"
ThisWorkbook.VBProject.VBComponents("F_Export").Export "D:\OneDrive - MSF\Documents\Source\F_Export.frm"
ThisWorkbook.VBProject.VBComponents("F_NomVisible").Export "D:\OneDrive - MSF\Documents\Source\F_NomVisible.frm"

ThisWorkbook.VBProject.VBComponents("M_CreationFeuille").Export "D:\OneDrive - MSF\Documents\Source\M_CreationFeuille.bas"
ThisWorkbook.VBProject.VBComponents("M_Dico").Export "D:\OneDrive - MSF\Documents\Source\M_Dico.bas"
ThisWorkbook.VBProject.VBComponents("M_Export").Export "D:\OneDrive - MSF\Documents\Source\M_Export.bas"
ThisWorkbook.VBProject.VBComponents("M_FonctionsDivers").Export "D:\OneDrive - MSF\Documents\Source\M_FonctionsDivers.bas"
ThisWorkbook.VBProject.VBComponents("M_FonctionsTransf").Export "D:\OneDrive - MSF\Documents\Source\M_FonctionsTransf.bas"
ThisWorkbook.VBProject.VBComponents("M_Geo").Export "D:\OneDrive - MSF\Documents\Source\M_Geo.bas"
ThisWorkbook.VBProject.VBComponents("M_LineList").Export "D:\OneDrive - MSF\Documents\Source\M_LineList.bas"
ThisWorkbook.VBProject.VBComponents("M_Main").Export "D:\OneDrive - MSF\Documents\Source\M_Main.bas"
ThisWorkbook.VBProject.VBComponents("M_Migration").Export "D:\OneDrive - MSF\Documents\Source\M_Migration.bas"
ThisWorkbook.VBProject.VBComponents("M_NomVisible").Export "D:\OneDrive - MSF\Documents\Source\M_NomVisible.bas"
ThisWorkbook.VBProject.VBComponents("M_Traduction").Export "D:\OneDrive - MSF\Documents\Source\M_Traduction.bas"
ThisWorkbook.VBProject.VBComponents("M_Validation").Export "D:\OneDrive - MSF\Documents\Source\M_Validation.bas"

End Sub
