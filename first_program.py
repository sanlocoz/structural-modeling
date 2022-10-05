# -*- coding: utf-8 -*-
"""
Created on Sun Jan  3 14:43:43 2021

@author: mikhsanikhsan99@gmail.com/Muhammad Ikhsan, Civil Engineering Department, Sriwijaya University, 2021

Program ini dibuat untuk pemodelan parsial tugas akhir berjudul:
ANALISIS PENGARUH BEBAN TSUNAMI DAN GEMPA TERHADAP GEDUNG BETON BERTULANG 10 LANTAI BERDASARKAN SNI 1727:2020 DENGAN VARIASI KETINGGIAN TSUNAMI
"""

import os
import sys
import comtypes.client
import pandas

#Line 18-62 retrieved from ETABS Documentation
#Set this to True if we are attaching API to existing Instance
AttachToInstance = True

#set this to True if we are going to specify ETABS path manually
SpecifyPath =False

#specify ETABS path here
ProgramPath = "C:\Program Files (x86)\Computers and Structures\ETABS 18\ETABS.exe"

#full path to model
APIPath = "D:\Coba API Ikhsan"
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass
ModelPath = APIPath + os.sep + "Ikhsan skripsi.edb" #File name

if AttachToInstance:
    try:
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject") 
    except (OSError, comtypes.COMError):
        print("Tidak ada program ETABS yang dijalankan.")
        sys.exit(-1)

else:
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    if SpecifyPath:
        try:
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Tidak bisa memulai ETABS dari " + ProgramPath)
            sys.exit(-1)
    else:
        try: 
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject") 
        except (OSError, comtypes.COMError):
            print("Tidak dapat memulai ETABS.")
            sys.exit(-1)

#create SapModel object
SapModel = myETABSObject.SapModel

N_mm_C = 9
kN_m_C = 6

file_input="Parameter Rencana Tambah.xlsx"

SapModel.SetPresentUnits(N_mm_C)
#concrete material properties
input_beton = pandas.read_excel(file_input,sheet_name=1,usecols="A:B")
nama_beton=[]
beton = 2
for i in range(input_beton.shape[0]):
    name="Beton "+str(input_beton.iat[i,1])+" MPa"
    nama_beton.append(name);
    ret = SapModel.PropMaterial.AddMaterial(name, beton, "Indonesia", "-", str(input_beton.iat[i,0])+" MPa",name)
    #variable to retrieve data from default properties
    E=0;U=0;A=0;G=0
    temp=[0]*10
    #retrieving data from default properties
    [E,U,A,G,reto] = SapModel.PropMaterial.GetMPIsotropic(name,E,U,A,G)
    temp = SapModel.PropMaterial.GetOConcrete_1(name,temp[0],temp[1],temp[2],temp[3],temp[4],temp[5],temp[6],temp[7],temp[8],temp[9])
    #changing data to user input
    temp[0]=input_beton.iat[i,1] #compressive strength
    E = 4700*(input_beton.iat[i,1])**0.5 #modulus of elasticity
    ret = SapModel.PropMaterial.SetMPIsotropic(name,E,U,A)
    ret = SapModel.PropMaterial.SetOConcrete_1(name,temp[0],temp[1],temp[2],temp[3],temp[4],temp[5],temp[6],temp[7],temp[8],temp[9])
#    print(ret)
    
#rebar material properties
input_tulangan = pandas.read_excel(file_input,sheet_name=2,usecols="A:C")
nama_tulangan=[]
tulangan = 6
for i in range(input_tulangan.shape[0]):
    name="Tulangan fy "+str(input_tulangan.iat[i,1])+" MPa"
    nama_tulangan.append(name);
    ret = SapModel.PropMaterial.AddMaterial(name, tulangan, "Indonesia", "-", "fy= "+str(input_tulangan.iat[i,0])+" MPa",name)
    #variable to retrieve data from default properties
    E=0;A=0
    temp=[0]*10
    #retrieving data from default properties
    [E,A,ret] = SapModel.PropMaterial.GetMPUniaxial(name,E,A)
    temp = SapModel.PropMaterial.GetORebar_1(name,temp[0],temp[1],temp[2],temp[3],temp[4],temp[5],temp[6],temp[7],temp[8],temp[9])
    #changing data to user input
    temp[0]=input_tulangan.iat[i,1] #fy
    temp[1]=input_tulangan.iat[i,2] #fu
    temp[2]=temp[0]*1.1 #Efy
    temp[3]=temp[1]*1.1 #Efu
    E = 200000 #modulus of elasticity
    ret = SapModel.PropMaterial.SetMPUniaxial(name,E,A)
    ret = SapModel.PropMaterial.SetORebar_1(name,temp[0],temp[1],temp[2],temp[3],temp[4],temp[5],temp[6],temp[7],temp[8],temp[9])
    print(ret)
    
#changing weight per unit volume concrete (2400 kgf/m3) rebar (7850 kgf/m3)
kgf_m_C=8
SapModel.SetPresentUnits(kgf_m_C)
for i in range(input_beton.shape[0]):
    name=nama_beton[i]
    ret = SapModel.PropMaterial.SetWeightAndMass(name,1,2400)
    
for i in range(input_tulangan.shape[0]):
    name=nama_tulangan[i]
    ret = SapModel.PropMaterial.SetWeightAndMass(name,1,7850)
    
#making rectangular beam section
input_balok = pandas.read_excel(file_input,sheet_name=3,usecols="A:F")
SapModel.SetPresentUnits(N_mm_C)

for i in range(input_beton.shape[0]):
    for j in range(input_balok.shape[0]):
        name="B "+nama_beton[i].replace("Beton ","")+" "+str(input_balok.iat[j,1])+" x "+str(input_balok.iat[j,2])
        ret = SapModel.PropFrame.SetRectangle(name, nama_beton[i],input_balok.iat[j,2],input_balok.iat[j,1])
        
        mod = [1, 1, 1, 0.15, 0.35, 0.35, 1, 1] #modifier for beams
        ret = SapModel.PropFrame.SetModifiers(name, mod)
        
        longit=input_balok.iat[j,3]-1 #longitudinal reinforcement material
        conf=input_balok.iat[j,4]-1 #confinement reinforcement material
        ret = SapModel.PropFrame.SetRebarBeam(name, nama_tulangan[longit], nama_tulangan[conf], input_balok.iat[j,5], input_balok.iat[j,5], 0, 0, 0, 0)
        print(ret)
        
#making rectangular column section
input_kolom = pandas.read_excel(file_input,sheet_name=4,usecols="A:N")
nama_kolom = []
for i in range(input_beton.shape[0]):
    for j in range(input_kolom.shape[0]):
        name="K "+nama_beton[i].replace("Beton ","")+" "+str(input_kolom.iat[j,1])+" x "+str(input_kolom.iat[j,2])
        nama_kolom.append(name)
        ret = SapModel.PropFrame.SetRectangle(name, nama_beton[i],input_kolom.iat[j,1],input_kolom.iat[j,2])
        
        mod = [1, 1, 1, 1, 0.7, 0.7, 1, 1] #modifier for columns
        ret = SapModel.PropFrame.SetModifiers(name, mod)
        
        longit=input_kolom.iat[j,3]-1 #longitudinal reinforcement material
        conf=input_kolom.iat[j,4]-1 #confinement reinforcement material
        cover=input_kolom.iat[j,5] #clear cover
        n3bars=input_kolom.iat[j,8] #number of longitudinal bars along 3 dir face
        n2bars=input_kolom.iat[j,9] #number of longitudinal bars along 2 dir face
        longsize=str(input_kolom.iat[j,6]) #longitudinal bar size
        tiesize=str(input_kolom.iat[j,7]) #tie bar size
        tiespace=input_kolom.iat[j,10] #tie bar spacing along 1 axis
        n2dirtie=input_kolom.iat[j,12] #number of confinement bar in 2 dir face
        n3dirtie=input_kolom.iat[j,11] #number of confinement bar in 3 dir face
        tobedesign=True if input_kolom.iat[j,13] == "Design" else False #design/check
        ret = SapModel.PropFrame.SetRebarColumn(name, nama_tulangan[longit], nama_tulangan[conf], 1, 1, cover, 0, n3bars, n2bars, longsize, tiesize, tiespace, n2dirtie, n3dirtie, tobedesign)
#        print(ret)
#making slab section
input_pelat = pandas.read_excel(file_input,sheet_name=5,usecols="A:B")

for i in range(input_beton.shape[0]):
    for j in range(input_pelat.shape[0]):
        name="P "+nama_beton[i].replace("Beton ","")+" "+str(input_pelat.iat[j,1])
        slab=0
        shell_thin=1
        
        ret = SapModel.PropArea.SetSlab(name, slab, shell_thin, nama_beton[i], input_pelat.iat[j,1])
        
        mod = [0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 1, 1, 1, 1] #modifier for slab
        ret = SapModel.PropArea.SetModifiers(name, mod)
