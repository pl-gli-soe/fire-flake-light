Attribute VB_Name = "RunFromMCDModule"
Public Sub sDUNS(mcd As CommonData, r As Range)

    r = mcd.duns

End Sub

Public Sub sSUPPLIER(mcd As CommonData, r As Range)

    r = mcd.supplierName

End Sub

Public Sub sF_U(mcd As CommonData, r As Range)
    r = mcd.fupCode
End Sub

Public Sub sA(mcd As CommonData, r As Range)
    r = mcd.fmaFupCode
End Sub

Public Sub sMISC(mcd As CommonData, r As Range)
    r = mcd.misc
End Sub

Public Sub sDOH(mcd As CommonData, r As Range)
    r = mcd.doh
End Sub

Public Sub sOS(mcd As CommonData, r As Range)
    r = mcd.os
End Sub


Public Sub sBANK(mcd As CommonData, r As Range)
    r = mcd.bank
End Sub

Public Sub sBBAL(mcd As CommonData, r As Range)
    r = mcd.bbal
End Sub

Public Sub sCBAL(mcd As CommonData, r As Range)
    r = mcd.cbal
End Sub

Public Sub sPCS_TO_GO(mcd As CommonData, r As Range)
    r = mcd.pcsToGo
End Sub

Public Sub sDK(mcd As CommonData, r As Range)
End Sub

Public Sub sMODE(mcd As CommonData, r As Range)
    r = mcd.mode
End Sub

Public Sub sMNPC(mcd As CommonData, r As Range)
    
End Sub

Public Sub sNCX(mcd As CommonData, r As Range)
    
End Sub

Public Sub sOBS(mcd As CommonData, r As Range)

End Sub

Public Sub sSTD_PACK(mcd As CommonData, r As Range)
    r = mcd.stdPack
End Sub

Public Sub soneJOB(mcd As CommonData, r As Range)

End Sub

Public Sub sIP(mcd As CommonData, r As Range)
End Sub

Public Sub sCOUNT(mcd As CommonData, r As Range)
    r = mcd.count_cmnt
End Sub

Public Sub sO(mcd As CommonData, r As Range)
    r = mcd.o_cmnt
End Sub

Public Sub sF(mcd As CommonData, r As Range)
    r = mcd.f_cmnt
End Sub

Public Sub sPART_NAME(mcd As CommonData, r As Range)
    r = mcd.partName
End Sub

Public Sub sQHD(mcd As CommonData, r As Range)
    r = mcd.qhd
End Sub

Public Sub sTT(mcd As CommonData, r As Range)
    r = mcd.ttime
End Sub
