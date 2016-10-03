VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = "client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInfo As clsClientinfo
Private mCA As clsClientCA

Public Property Set Info(objInfo As clsClientinfo)
    Set mInfo = objInfo
End Property

Public Property Get Info() as clsClientinfo
    Set Info = mInfo
End Property

Public Property Set CA(objCA As clsClientCA)
    Set mCA = objCA
End Property

Public Property Get CA() as clsClientCA
    Set CA = mCA
End Property
    
