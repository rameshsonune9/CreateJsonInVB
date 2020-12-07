VERSION 5.00
Begin VB.Form frmInitialize 
   Caption         =   "Initialize"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim jsonItems As New Collection
Dim jsonDictionary As New Scripting.Dictionary
Dim jsonFileObject As New Scripting.FileSystemObject
Dim jsonFileExport As TextStream


Set cMain = New Collection

Set dMain = New Dictionary
dMain("Version") = "1.1"

Set d0 = New Dictionary
d0("TaxSch") = "GST"
d0("SupTyp") = "B2B"
d0("IgstOnIntra") = "N"
d0("RegRev") = "N"
d0("EcmGstin") = Null
dMain.Add "TranDtls", d0

Set d1 = New Dictionary
d1("Typ") = "INV"
d1("No") = "NICE/CTD1"
d1("Dt") = "06/08/2020"
dMain.Add "DocDtls", d1

Set d2 = New Dictionary
d2("Gstin") = ""
d2("LglNm") = ""
d2("TrdNm") = ""
d2("Addr1") = ""
d2("Addr2") = ""
d2("Loc") = ""
d2("Pin") = 0
d2("Stcd") = ""
d2("Ph") = ""
d2("Em") = ""
dMain.Add "SellerDtls", d2

Set d3 = New Dictionary
d3("Gstin") = ""
d3("LglNm") = ""
d3("TrdNm") = ""
d3("Pos") = "3"
d3("Addr1") = ""
d3("Addr2") = ""
d3("Loc") = ""
d3("Pin") = 0
d3("Stcd") = ""
d3("Ph") = ""
d3("Em") = ""
dMain.Add "BuyerDtls", d3

Set d4 = New Dictionary
d4("Nm") = ""
d4("Addr1") = ""
d4("Addr2") = ""
d4("Loc") = ""
d4("Pin") = 0
d4("Stcd") = ""
dMain.Add "DispDtls", d4

Set d5 = New Dictionary
d5("Gstin") = ""
d5("LglNm") = ""
d5("TrdNm") = ""
d5("Addr1") = ""
d5("Addr2") = ""
d5("Loc") = ""
d5("Pin") = 0
d5("Stcd") = ""
dMain.Add "ShipDtls", d5

Set d6 = New Dictionary
d6("AssVal") = 0
d6("IgstVal") = 0
d6("CgstVal") = 0
d6("SgstVal") = 0
d6("CesVal") = 0
d6("StCesVal") = 0
d6("Discount") = 0
d6("OthChrg") = 0
d6("RndOffAmt") = 0
d6("TotInvVal") = 0
d6("TotInvValFc") = 0
dMain.Add "ValDtls", d6

dMain("ExpDtls") = Null

Set dewb = New Dictionary
dewb("TransId") = ""
dewb("TransName") = ""
dewb("TransMode") = ""
dewb("Distance") = 0
dewb("TransDocNo") = ""
dewb("TransDocDt") = ""
dewb("VehNo") = ""
dewb("VehType") = ""
dMain.Add "EwbDtls", dewb
'dMain("EwbDtls") = Null

dMain("PayDtls") = Null
dMain("RefDtls") = Null

Set d7 = New Dictionary
d7("Url") = Null
d7("Docs") = Null
d7("Info") = Null
dMain.Add "AddlDocDtls", d7

Set cItem = New Collection
For r = 0 To 2
    Set dItem = New Dictionary
    dItem("SlNo") = r + 1
    dItem("PrdDesc") = "steel"
    dItem("IsServc") = "N"
    dItem("HsnCd") = ""
    dItem("Barcde") = Null
    dItem("Qty") = 0
    dItem("FreeQty") = 0
    dItem("Unit") = ""
    dItem("UnitPrice") = 0
    dItem("TotAmt") = 0
    dItem("Discount") = 0
    dItem("PreTaxVal") = 0
    dItem("AssAmt") = 0
    dItem("GstRt") = 0
    dItem("IgstAmt") = 0
    dItem("CgstAmt") = 0
    dItem("SgstAmt") = 0
    dItem("CesRt") = 0
    dItem("CesAmt") = 0
    dItem("CesNonAdvlAmt") = 0
    dItem("StateCesRt") = 0
    dItem("StateCesAmt") = 0
    dItem("StateCesNonAdvlAmt") = 0
    dItem("OthChrg") = 0
    dItem("TotItemVal") = 0
    dItem("OrdLineRef") = Null
    dItem("OrgCntry") = Null
    dItem("PrdSlNo") = Null
    
    Set d8 = New Dictionary
    d8("Nm") = ""
    d8("ExpDt") = ""
    d8("WrDt") = ""
    dItem.Add "BchDtls", d8

    Set cd9 = New Collection
    Set d9 = New Dictionary
    d9("Nm") = ""
    d9("Val") = Null
    cd9.Add d9
    dItem.Add "AttribDtls", cd9
    
    cItem.Add dItem
Next
dMain.Add "ItemList", cItem

cMain.Add dMain

Set jsonFileExport = jsonFileObject.CreateTextFile("C:\Ramesh\VB6 New\ICSEInvoice\jsonExample.json", True)
'jsonFileExport.WriteLine (JsonConverter.ConvertToJson(c1, Whitespace:=3))

TempTxt = JsonConverter.ConvertToJson(cMain, Whitespace:=3)
jsonFileExport.WriteLine TempTxt
Debug.Print TempTxt
End Sub

