# -*- coding: utf-8 -*-
"""
Created on Mon Mar  8 21:51:55 2021

@author: user
"""

import os
import sys
import comtypes.client

AttachToInstance=True 

if AttachToInstance:
    try:
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("Tidak ada program ETABS yang dijalankan.")
        sys.exit(-1)

#create SapModel object
SapModel = myETABSObject.SapModel

kN_m_C = 6
SapModel.SetPresentUnits(kN_m_C)

reto=SapModel.LoadCases.ResponseSpectrum.GetLoads("RSX0")
SF_now_x=reto[3][0]
retoy=SapModel.LoadCases.ResponseSpectrum.GetLoads("RSY0")
SF_now_y=retoy[3][0]

ret = SapModel.Analyze.RunAnalysis()

ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

ret = SapModel.Results.Setup.SetCaseSelectedForOutput("RSX0")
ret = SapModel.Results.Setup.SetCaseSelectedForOutput("RSY0")
ret = SapModel.Results.Setup.SetCaseSelectedForOutput("EQx")
ret = SapModel.Results.Setup.SetCaseSelectedForOutput("EQy")

NumberResults=0
LoadCase=[]
StepType=[]
StepNum=[]
Fx=[]
Fy=[]
Fz=[]
Mx=[]
ParamMy=[]
Mz=[]
gx=0
gy=0
gz=0
[NumberResults, LoadCase, StepType, StepNum, Fx, Fy, Fz, Mx, ParamMy, Mz, gx, gy, gz,ret]= SapModel.Results.BaseReact(NumberResults, LoadCase, StepType, StepNum, Fx, Fy, Fz, Mx, ParamMy, Mz, gx, gy, gz)

SFx=abs(Fx[0])/abs(Fx[2])
SFy=abs(Fy[1])/abs(Fy[3])
SFmin=9.81/8

SFx*=SF_now_x
SFy*=SF_now_y

SFx=max(SFmin,SFx)
SFy=max(SFmin,SFy)

reti = SapModel.SetModelIsLocked(False)

#ret=SapModel.RespCombo.SetCaseList("RSX0",0,"U1",SFx)
ret=SapModel.LoadCases.ResponseSpectrum.SetLoads("RSX0",reto[0],reto[1],reto[2],(SFx,),reto[4],reto[5])
ret=SapModel.LoadCases.ResponseSpectrum.SetLoads("RSX1",reto[0],reto[1],reto[2],(SFx,),reto[4],reto[5])
ret=SapModel.LoadCases.ResponseSpectrum.SetLoads("RSY0",retoy[0],retoy[1],retoy[2],(SFy,),retoy[4],retoy[5])
ret=SapModel.LoadCases.ResponseSpectrum.SetLoads("RSY1",retoy[0],retoy[1],retoy[2],(SFy,),retoy[4],retoy[5])

ret = SapModel.Analyze.RunAnalysis()