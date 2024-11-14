import os
import comtypes.client
import pandas as pd
import math
import numpy as np
#------------------------------------------------------------------------------------------
#Inicialización de varibales globales
nom_conc = None
nom_ac = None
iter=1
P_t=0
T_obj=0
#------------------------------------------------------------------------------------------
#Simplificaciones del modelo
w_m=0.25
w_m_azo=0.1
w_sc_azo=0.1
w_sc=[0.3,0.2,0.4]  #sala de operaciones y zona de servicio, cuartos, corredores y escaleras
#------------------------------------------------------------------------------------------
def initial_etabs():
    #create API helper object
    helper = comtypes.client.CreateObject('ETABSv1.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
    #create an instance of the ETABS object from the latest installed ETABS
    myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject") 
    #start ETABS application
    myETABSObject.ApplicationStart()
    SapModel  = myETABSObject.SapModel
    #initialize model
    SapModel.InitializeNewModel(12)
    return SapModel 

def sup_cel(num,precision):
    rpta=math.ceil(num/precision)*precision
    return rpta

def def_secc(SapModel,nom,mat_conc,mat_ac,lados):
    ret = SapModel.PropFrame.SetRectangle(nom, mat_conc, lados[0], lados[1])
    if nom[0]=='V':
        ret = SapModel.PropFrame.SetModifiers(nom, [1,1,1,0.001,1,1,1,1])
        ret = SapModel.PropFrame.SetRebarBeam(nom, mat_ac, mat_ac, 0.08, 0.08, 0, 0, 0, 0)
    elif nom[0]=='C':
        TableKey='Frame Section Property Definitions - Concrete Column Reinforcing'
        TableVersion = 1
        FieldsKeysIncluded=[
        'Name','RebarMatL','RebarMatC','ReinfConfig','IsSpiral','IsDesigned',
        'Cover','NumBars3Dir','NumBars2Dir','NumBarsCirc',
        'BarSizeLong','BarSizeCorn','BarSizeConf','SpacingConf','NumCBars3','NumCBars2',
        ]
        NumberRecords=1
        TableData=[
            nom, mat_ac, mat_ac, 'Rectangular', 'Yes', 'Yes',
            '0.06', '4', '4', '0',
            '20', '20', '10', '0.15', '3', '3'
        ]
        ret = SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
        FillImport = True
        ret= SapModel.DatabaseTables.ApplyEditedTables(FillImport) 

def const_sis(z_sis,s_sis):
    """
    This function creates a frame element using two coordenates, the section property and the name of the element

    Args:
        z_sis (int): value of the zone factor
        s_sis (int): value of the soil factor
    Returns:
        cons_z (double): the constant of the zone factor
        cons_s (double): the constant of the soil factor
        cons_tp (double): first constant of the seismic graph
        cons_tp (double): second constant of the seismic graph
    """
    tabla_z=pd.DataFrame({'Z':[0.45,0.35,0.25,0.1]},index=[4,3,2,1])
    tabla_s=pd.DataFrame({0:[0.8]*4,1:[1]*4,2:[1.05,1.15,1.2,1.6],3:[1.1,1.2,1.4,2]},index=[4,3,2,1])
    tabla_t=pd.DataFrame({0:[0.3,3],1:[0.4,2.5],2:[0.6,2],3:[1,1.6]},index=['Tp','Tl'])

    cons_z=tabla_z.loc[z_sis,'Z']
    cons_s=tabla_s.loc[z_sis,s_sis]
    cons_tp=tabla_t.loc['Tp',s_sis]
    cons_tl=tabla_t.loc['Tl',s_sis]

    return [cons_z,cons_s,cons_tp,cons_tl]

def calcular_C(T_edi, Tp, Tl, norma):
    if norma=='030':
        cte=0
    else:
        cte=1

    if T_edi<Tp*0.2*cte:
        C = 1+7.5*(T_edi/Tp)
    elif T_edi<Tp:
        C = 2.5
    elif Tp <= T_edi<Tl:
        C = 2.5*(Tp/T_edi)
    else:  # T_edi >= Tl
        C = 2.5*(Tp*Tl/T_edi**2)
    return C

def espectro_func(z_sis,u_sis,s_sis,r_sis,norma):
    [cons_z,cons_s,cons_tp,cons_tl]=const_sis(z_sis,s_sis)

    if norma=='030':
        cte=1
    else:
        cte=1.5
        u_sis=1
        r_sis=1

    T=[0]*501
    C=[0]*501
    c_val=pd.DataFrame({'T':T,'C':C})

    for i in range(501):
        T[i]=i*0.02
        C[i]=calcular_C(T[i],cons_tp,cons_tl,norma)

    c_val=pd.DataFrame({'T':T,'sa':C})
    c_val['sa']=cte*c_val['sa']*cons_z*u_sis*cons_s/r_sis
    return c_val

def define_espectrum(SapModel,nom,norma):
    TableKey='Functions - Response Spectrum - User Defined'
    TableVersion = 1
    FieldsKeysIncluded=[
    'Name',
    'Period',
    'Value',
    'DampRatio',
    'GUID',
    ]
    espectro=espectro_func(Z_sis,U_sis,S_sis,R_sis,norma)
    list1 = []
    # Iterate over rows and build the string
    for i in range(1, espectro.shape[0]):
        if i == 1:
            list1+=[nom,str(espectro["T"].iloc[i-1]),str(espectro["sa"].iloc[i-1]), '0.05', '',]
        else:
            list1+=[nom,str(espectro["T"].iloc[i-1]),str(espectro["sa"].iloc[i-1]), '', '',]
    NumberRecords=1
    TableData=list1
    ret = SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
    FillImport = True
    ret= SapModel.DatabaseTables.ApplyEditedTables(FillImport) 

#Define modals
def def_loadcases(SapModel,name,type,arguments):
    if type=='StaticLinear':
        ret = SapModel.LoadCases.StaticLinear.SetCase(name)
        ret = SapModel.LoadCases.StaticLinear.SetLoads(name, arguments[0], arguments[1], arguments[2], arguments[3])
    elif type=='ModalEigen':
        ret = SapModel.LoadCases.ModalEigen.SetCase(name)
        ret = SapModel.LoadCases.ModalEigen.SetNumberModes(name, arguments[0], arguments[1])
        if len(arguments) > 2:  
            ret = SapModel.LoadCases.ModalEigen.SetInitialCase(name, arguments[2])
    elif type=='StaticNonlinear':
        ret = SapModel.LoadCases.StaticNonlinear.SetCase(name)
        ret = SapModel.LoadCases.StaticNonlinear.SetMassSource(name,arguments[0])
    elif type=='ResponseSpectrum':
        ret = SapModel.LoadCases.ResponseSpectrum.SetCase(name)
        ret = SapModel.LoadCases.ResponseSpectrum.SetLoads(name,arguments[0],arguments[1],arguments[2],arguments[3],arguments[4],arguments[5])
        ret = SapModel.LoadCases.ResponseSpectrum.SetModalCase(name,arguments[6])

def conf_etabs(SapModel,nom_esp):
    #define material property
    global nom_conc
    nom_conc=f'concreto {fc_concreto_p1} kg/cm2'
    ret = SapModel.PropMaterial.SetMaterial(nom_conc, 2)
    ret = SapModel.PropMaterial.SetMPIsotropic(nom_conc, (15100*fc_concreto_p1**0.5)*0.10160469053143*100, 0.15, 9.90E-06)
    ret = SapModel.PropMaterial.SetWeightAndMass(nom_conc, 1, 2.400)
    ret = SapModel.PropMaterial.SetOConcrete_1(nom_conc, (fc_concreto_p1)*9.8420653098757, False, 0, 1, 4, 0.0022, 0.0052, -0.1, 0, 0)
    global nom_ac
    nom_ac=f'acero {fy_acero_p1} kg/cm2'
    ret = SapModel.PropMaterial.SetMaterial(nom_ac, 6)
    ret = SapModel.PropMaterial.SetWeightAndMass(nom_ac, 1, 7.850)
    ret = SapModel.PropMaterial.SetORebar_1(nom_ac, fy_acero_p1*9.8420653098757, fy_acero_p1*9.8420653098757, 2*10**10, 0, 1, 0, 0.02, 0.1, 0, False)

    #define rectangular frame section property
    direct_secc={'C1':[l_c1]*2,
                'C2':[bl_vigac[0]+0.1]*2,
                'V1':bl_viga,
                'V2':bl_vigac}         #L/8 para vigas de base

    for key in direct_secc:
        def_secc(SapModel,key,nom_conc,nom_ac,direct_secc[key])

    #define slab section property
    ret = SapModel.PropArea.SetSlab("L1", 0, 3, nom_conc, e_losa)
    ret = SapModel.PropArea.SetModifiers("L1", [1,1,1,1,1,1,1,1,1,1])

    #define combo cases
    direct_load={'CM':1,'CV':3,'CSx':5,'CSy':5}
    for key in direct_load:
        ret = SapModel.LoadPatterns.Add(key,direct_load[key])
    ret = SapModel.LoadPatterns.SetSelfWTMultiplier("CM", 1)

    #define Mass source
    TableKey='Mass Source Definition'
    TableVersion = 1
    FieldsKeysIncluded=['Name','IsDefault','IncLateral','IncVertical','LumpMass','SourceSelf','SourceAdded',
                        'SourceLoads','MoveMass','RatioX','RatioY','LoadPattern','Multiplier','GUID']
    NumberRecords=1
    direct_mass={'MsSrc1':['No','0','0',['CM','CV'],['1','0.5'],''],
                'MasaX-':['Yes','-0.05','0',['CM','CV'],['1','0.5'],''],
                'MasaX+':['Yes','0.05','0',['CM','CV'],['1','0.5'],''],
                'MasaY-':['Yes','0','-0.05',['CM','CV'],['1','0.5'],''],
                'MasaY+':['Yes','0','0.05',['CM','CV'],['1','0.5'],''],
                }
    # Iterate over rows and build the string
    TableData=[]
    for key in direct_mass:
        for i in range(1,len(direct_mass[key][3])+1):
            TableData+=[key,'No','Yes','No','Yes','No','No','Yes']+direct_mass[key][0:3]+[direct_mass[key][3][i-1],direct_mass[key][4][i-1],'',]

    ret = SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
    FillImport = True
    ret= SapModel.DatabaseTables.ApplyEditedTables(FillImport)

    mod_d=3*Num_Pisos

    if Disty<Distx:
        coef_X=0.3
        coef_Y=1
    else:
        coef_X=1
        coef_Y=0.3
    
    direct_loadcases={
        "Peso Sísmico":['StaticLinear',[2, ['Load']*2, ['CM','CV'], [1,0.5]]],
        "Modal":['ModalEigen',[mod_d,1]],
        "MasaY+":['StaticNonlinear',['MasaY+']],
        "MasaY-":['StaticNonlinear',['MasaY-']],
        "MasaX+":['StaticNonlinear',['MasaX+']],
        "MasaX-":['StaticNonlinear',['MasaX-']],
        "ModalMasaY+":['ModalEigen',[mod_d,1,"MasaY+"]],
        "ModalMasaY-":['ModalEigen',[mod_d,1,"MasaY-"]],
        "ModalMasaX+":['ModalEigen',[mod_d,1,"MasaX+"]],
        "ModalMasaX-":['ModalEigen',[mod_d,1,"MasaX-"]],
        "ESPECTRAL Y-Y":['ResponseSpectrum',[1,['U2'],[nom_esp],[9.8067],['Global'],[0],'Modal']],
        "ESPECTRAL X-X":['ResponseSpectrum',[1,['U1'],[nom_esp],[9.8067],['Global'],[0],'Modal']],
        "ESP X-X MY+":['ResponseSpectrum',[2,['U1','U2'],[nom_esp]*2,[9.8067*coef_X,9.8067*coef_Y],['Global']*2,[0]*2,'ModalMasaY+']],
        "ESP X-X MY-":['ResponseSpectrum',[2,['U1','U2'],[nom_esp]*2,[9.8067*coef_X,9.8067*coef_Y],['Global']*2,[0]*2,'ModalMasaY-']],
        "ESP Y-Y MX+":['ResponseSpectrum',[2,['U2','U1'],[nom_esp]*2,[9.8067*coef_X,9.8067*coef_Y],['Global']*2,[0]*2,'ModalMasaX+']],
        "ESP Y-Y MX-":['ResponseSpectrum',[2,['U2','U1'],[nom_esp]*2,[9.8067*coef_X,9.8067*coef_Y],['Global']*2,[0]*2,'ModalMasaX-']]
    }
    for key in direct_loadcases:
        def_loadcases(SapModel,key,direct_loadcases[key][0],direct_loadcases[key][1])

    ret = SapModel.RespCombo.Add('Sx', 1)
    ret = SapModel.RespCombo.SetCaseList('Sx', 0, "ESP X-X MY+", 1)
    ret = SapModel.RespCombo.SetCaseList('Sx', 0, "ESP X-X MY-", 1)
    ret = SapModel.RespCombo.Add('Sy', 1)
    ret = SapModel.RespCombo.SetCaseList('Sy', 0, "ESP Y-Y MX+", 1)
    ret = SapModel.RespCombo.SetCaseList('Sy', 0, "ESP Y-Y MX-", 1)

    comb=pd.DataFrame({'Combinaciones':['CM','CV','Sx','Sy'],
                    'CM+CV':[1,1,0,0],
                    '1.4CM+1.7CV':[1.4,1.7,0,0],
                    '1.25(CM+CV)+CSX':[1.25,1.25,1,0],
                    '1.25(CM+CV)-CSX':[1.25,1.25,-1,0],
                    '1.25(CM+CV)+CSY':[1.25,1.25,0,1],
                    '1.25(CM+CV)-CSY':[1.25,1.25,0,-1],
                    '0.9CM+CSX':[0.9,0,1,0],
                    '0.9CM-CSX':[0.9,0,-1,0],
                    '0.9CM+CSY':[0.9,0,0,1],
                    '0.9CM-CSY':[0.9,0,0,-1]})

    direct_casecomb={'CM':0,'CV':0,'Sx':1,'Sy':1}

    for i in range(1, comb.shape[1]):
        ret = SapModel.RespCombo.Add(comb.columns[i], 0)
        for j in range(comb.shape[0]):
            ret = SapModel.RespCombo.SetCaseList(comb.columns[i], direct_casecomb[comb['Combinaciones'][j]], comb['Combinaciones'][j], comb.loc[j, comb.columns[i]])

    #set diaphragm
    ret = SapModel.Diaphragm.SetDiaphragm("D1", False)
    return SapModel

def def_GridSystem(SapModel,df_coordsx,df_coordsy):
    TableKey='Grid Definitions - Grid Lines'
    TableVersion = 1
    FieldsKeysIncluded=[
    'Name',
    'LineType',
    'ID',
    'Ordinate',
    'BubbleLoc',
    'Visible'
    ]
    NumberRecords=1
    TableData=[]
    
    grids_x=list(df_coordsx['Coords X'])
    grids_y=list(df_coordsy['Coords Y'])

    for i in range(0,len(grids_x)):
        TableData+=['G1','X (Cartesian)','X'+str(i+1),str(grids_x[i]),'End','Yes']
    for j in range(0,len(grids_y)):
        TableData+=['G1','Y (Cartesian)','Y'+str(j+1),str(grids_y[j]),'End','Yes']

    ret = SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
    ret= SapModel.DatabaseTables.ApplyEditedTables(True) 

def obtencion_db(SapModel,title_table):
    TableKey = title_table
    TableVersion = 1
    NumberFields = 0
    FieldKey = []
    FieldName = []
    Description = []
    UnitsString = []
    IsImportable=[]
    fields_table = SapModel.DatabaseTables.GetAllFieldsInTable(
        TableKey,
        TableVersion,
        NumberFields,
        FieldKey,
        FieldName,
        Description,
        UnitsString,
        IsImportable
    )
    GroupName=''
    TableVersion=1
    FieldsKeysIncluded=[]
    NumberRecords=0
    TableData=[]

    table = SapModel.DatabaseTables.GetTableForDisplayArray(TableKey, list(fields_table[2]), GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData) 
    vals = np.array_split(table[4],table[3])
    df = pd.DataFrame(vals)
    df.columns = table[2]
    return df

def viga_etabs(SapModel,coord_1, coord_2, sec, nom_prop):
    """
    This function creates a frame element using two coordenates, the section property and the name of the element

    Args:
        coord_1 (list, dim 3): Coordenate of the start point
        coord_2 (list, dim 3): Coordenate of the final point
        sec (string): Name of the section property
        nom_prop (string): Name of the element

    Returns:
        SapModel: A frame in the data of the etabs file
    """
    if len(coord_1)!=3:
        print('La coordenada debe ser de 3 dimensiones')
    else:
        ret = SapModel.FrameObj.AddByCoord(coord_1[0], coord_1[1], coord_1[2], coord_2[0], coord_2[1], coord_1[2], '', sec, nom_prop, 'Global')
        ret = SapModel.FrameObj.SetEndLengthOffset(nom_prop, True,0,0,0)
        ret = SapModel.FrameObj.SetInsertionPoint(nom_prop, 8, False, True, [0,0,0], [0,0,0])

def planta_BF(SapModel,numx,numy,Espaciamientox,Espaciamientoy,coord_x,coord_y,vigas_TF,cols_TF,carga_sc):
    global iter
    for k in range(0, Num_Pisos + 1):
        for i in range(1, numx + 1):
            for j in range(1, numy + 1):
                #Nombre de elementos
                name=str(iter)+'_'+str(i)+'_'+str(j)+'_'+str(k)
                #Coordenadas
                x_coord = (i - 1)*Espaciamientox+coord_x
                y_coord = (j - 1)*Espaciamientoy+coord_y
                if k == 1:
                    z_coord1 = 0
                    z_coord2 = Altura_piso1
                else:
                    z_coord1 = Altura_piso1 + (k - 2)*Altura_piso_tipico
                    z_coord2 = Altura_piso1 + (k - 1)*Altura_piso_tipico
                #Creación de Columnas
                if cols_TF and k>=1:
                    ret = SapModel.FrameObj.AddByCoord(x_coord, y_coord, z_coord1, x_coord, y_coord, z_coord2, 'z'+name, 'C1',  'Cz'+name, 'Global')                    
                    if k==1:
                        [PointName1, PointName2, ret] = SapModel.FrameObj.GetPoints('Cz'+name, '', '')
                        ret = SapModel.PointObj.SetRestraint(PointName1, [True]*6)
                #Creación de vigas
                if vigas_TF:                    
                    if k==0:
                        sec_v='V2'
                        z_coord_v=0
                    else:
                        z_coord_v=z_coord2
                        sec_v='V1'

                    if i!=1:
                        viga_etabs(SapModel,[x_coord,y_coord,z_coord_v],[x_coord-Espaciamientox,y_coord,z_coord_v],sec_v,sec_v+'x'+name)
                    if j!=1:
                        viga_etabs(SapModel,[x_coord,y_coord,z_coord_v],[x_coord,y_coord-Espaciamientoy,z_coord_v],sec_v,sec_v+'y'+name)
                #Creación de Losas
                if i!=1 and j!=1 and numx>1 and numy>1:
                    coord=[k,i,j]
                    x_losa=[x_coord,x_coord,x_coord-Espaciamientox,x_coord-Espaciamientox]
                    y_losa=[y_coord,y_coord-Espaciamientoy,y_coord-Espaciamientoy,y_coord]

                    if k==0:
                        name_losa='Lb'+name
                        z_coord_losa=0
                    else:
                        name_losa='L'+name
                        z_coord_losa=z_coord2

                    if k==Num_Pisos:
                        cargam_losa=w_m_azo
                        cargav_losa=w_sc_azo
                    else:
                        cargam_losa=w_m
                        cargav_losa=w_sc[carga_sc-1]

                    ret = SapModel.AreaObj.AddByCoord(4, x_losa, y_losa, [z_coord_losa]*4, name,'',name_losa)
                    ret = SapModel.AreaObj.SetProperty(name_losa, "L1",0)
                    ret = SapModel.AreaObj.SetDiaphragm(name_losa, "D1")

                    ret = SapModel.AreaObj.SetLoadUniform(name_losa, "CM", cargam_losa, 10, True)
                    ret = SapModel.AreaObj.SetLoadUniform(name_losa, "CV", cargav_losa, 10, True)
    iter+=1

def planta_BA(SapModel1,numx,numy,Espaciamientox,Espaciamientoy,coord_x,coord_y,vigas_TF,cols_TF,carga_sc):
    global iter
    for k in range(1, Num_Pisos + 3):
        for i in range(1, numx + 1):
            for j in range(1, numy + 1):
                #Nombre de elementos
                name=str(iter)+'_'+str(i)+'_'+str(j)+'_'+str(k)      
                #Coordenadas       
                x_coord = (i - 1)*Espaciamientox+coord_x
                y_coord = (j - 1)*Espaciamientoy+coord_y
                if k==1:
                    z_coord1 = 0
                    z_coord2 = Altura_link
                elif k==2:
                    z_coord1 = Altura_link 
                    z_coord2 = Altura_link+Altura_Base
                elif k==3:
                    z_coord1 = Altura_link+Altura_Base 
                    z_coord2 = Altura_link+Altura_Base + Altura_piso1
                else:
                    z_coord1 = Altura_link+Altura_Base + Altura_piso1 + (k - 4)*Altura_piso_tipico
                    z_coord2 = Altura_link+Altura_Base + Altura_piso1 + (k - 3)*Altura_piso_tipico
                #Creación de Columnas
                if cols_TF:
                    if k==1:
                        ret = SapModel1.LinkObj.AddByCoord(x_coord, y_coord, z_coord1, x_coord, y_coord, z_coord2, 'z'+name, False, 'Aislador01','Ca'+name, 'Global')
                        [PointName1, PointName2, ret] = SapModel1.LinkObj.GetPoints('Ca'+name, '', '')
                        ret = SapModel1.PointObj.SetRestraint(PointName1, [True]*6)
                    elif k==2:
                        ret = SapModel1.FrameObj.AddByCoord(x_coord, y_coord, z_coord1, x_coord, y_coord, z_coord2, 'z'+name, 'C2',  'Cz'+name, 'Global')
                    else:
                        ret = SapModel1.FrameObj.AddByCoord(x_coord, y_coord, z_coord1, x_coord, y_coord, z_coord2, 'z'+name, 'C1',  'Cz'+name, 'Global')
                #Creación de vigas
                if vigas_TF and k>=2:                    
                    if k==0:
                        sec_v='V2'
                    else:
                        sec_v='V1'

                    if i!=1:
                        viga_etabs(SapModel1,[x_coord,y_coord,z_coord2],[x_coord-Espaciamientox,y_coord,z_coord2],sec_v,sec_v+'x'+name)
                    if j!=1:
                        viga_etabs(SapModel1,[x_coord,y_coord,z_coord2],[x_coord,y_coord-Espaciamientoy,z_coord2],sec_v,sec_v+'y'+name)
                #Creación de Losas
                if i!=1 and j!=1 and k>=2:
                    coord=[k,i,j]
                    x_losa=[x_coord,x_coord,x_coord-Espaciamientox,x_coord-Espaciamientox]
                    y_losa=[y_coord,y_coord-Espaciamientoy,y_coord-Espaciamientoy,y_coord]

                    if k==2:
                        name_losa='Lb'+name
                    else:
                        name_losa='L'+name

                    if k==Num_Pisos+2:
                        cargam_losa=w_m_azo
                        cargav_losa=w_sc_azo
                    else:
                        cargam_losa=w_m
                        cargav_losa=w_sc[carga_sc-1]

                    ret = SapModel1.AreaObj.AddByCoord(4, x_losa, y_losa, [z_coord2]*4, name,'',name_losa)
                    ret = SapModel1.AreaObj.SetProperty(name_losa, "L1",0)
                    ret = SapModel1.AreaObj.SetDiaphragm(name_losa, "D1")

                    ret = SapModel1.AreaObj.SetLoadUniform(name_losa, "CM", cargam_losa, 10, True)
                    ret = SapModel1.AreaObj.SetLoadUniform(name_losa, "CV", cargav_losa, 10, True)                  
    iter+=1

def KC_calc(SapModel,amor_i,coef_tb):
    global P_t
    global T_obj
    #Periodos
    ##Obtención de data
    df=obtencion_db(SapModel,'Modal Periods And Frequencies')
    df=df[df['Case']=='Modal']
    periodo_1=float(df['Period'].iloc[0])

    #Masa
    ##Obtención de data
    df_masa=obtencion_db(SapModel,'Mass Summary by Story')
    ##Creación de la base de datos
    df_masa['Masa']=df_masa['UX'].astype(float)
    df_masa=df_masa.drop(['UX','UY','UZ'],axis=1)
    df_masa['Masa acumulada']=df_masa['Masa'].cumsum()
    P_t=df_masa['Masa acumulada'].max()*9.806

    #Extracted data from Rigid Base Model
    T_obj=coef_tb*periodo_1
    K_ten=(P_t/T_obj**2*4*np.pi**2/9.806)/(ctd_cols)

    C_cr=2*(ctd_cols*P_t*K_ten/9.806)**0.5
    C_edif=C_cr*amor_i
    C_disp=C_edif/(ctd_cols)
    
    print(df_masa)
    
    return [K_ten,C_disp]

def deriva_ines_max(SapModel):
    df=obtencion_db(SapModel,'Joint Drifts')
    df = df[df['OutputCase'].isin(['ESP X-X MY+','ESP X-X MY-','ESP Y-Y MX+','ESP Y-Y MX-',
                                'ESPECTRAL Y-Y E030','ESPECTRAL X-X E030'])]
    df[['DriftX','DriftY']]=df[['DriftX','DriftY']].astype(float)
    df = df[~df['Story'].isin(['Plataforma', 'Ref1'])]
    df['DriftG']=np.sqrt(df['DriftX'].fillna(0)**2 + df['DriftY'].fillna(0)**2)
    deriva_ines_max=df['DriftG'].max()
    return deriva_ines_max*1000

def calcular_T(T_edi):
    if T_edi<0.5:
        T = 1
    elif T_edi <= 2.5:
        T = 0.75+0.5*T_edi
    else:  # T_edi >= Tl
        T = 2
    return T
#------------------------------------------------------------------------------------------
os.system('cls')
#------------------------------------------------------------------------------------------
#draw
Num_Pisos =7
Altura_link=0.3
Altura_Base=1
Altura_piso1 =4.2
Altura_piso_tipico =3.2
fc_concreto_p1=280
fy_acero_p1=4200
Z_sis=4 # Factor de zona sísmica
U_sis=1.5 # Factor de uso de Base Fija (edificio esencial [necesita aisladores])
S_sis=1 # Tipo de suelo
R_sis=8 # Factor de reducción de Base Fija 

#dataframe para Grid
url_DatosPlanta=r'C:\Users\migue\OneDrive\Desktop\Datos_planta.xlsx'
df=pd.read_excel(url_DatosPlanta,index_col=0)
[df_largo,df_ancho]=df.shape
df[['numx','numy','carga_sc']]=df[['numx','numy','carga_sc']].astype(int)
df[['vigas_TF','cols_TF']]=df[['vigas_TF','cols_TF']].astype(bool)

#redefine Grid System
df_coordsx = pd.DataFrame()
df_coordsy = pd.DataFrame()
df_coordsGLB=pd.DataFrame()

for i in range(df_largo):
    coords_x = []
    coords_y = []

    # Extraer los primeros 6 valores de la lista de df_planta[key]
    [numx, numy, esp_x, esp_y, coordi_x, coordi_y] = list(df.iloc[i,0:6])
    cols_TF=df.iloc[i,7]

    # Generar las coordenadas x
    for i in range(numx):
        coords_x.append(i * esp_x + coordi_x)

    # Generar las coordenadas y
    for j in range(numy):
        coords_y.append(j * esp_y + coordi_y)

    # Asignar las coordenadas generadas a los DataFrames como una columna
    coords_x = [round(flt_x, 3) for flt_x in coords_x]
    coords_y = [round(flt_y, 3) for flt_y in coords_y]
    df_coordsx = pd.concat([df_coordsx, pd.DataFrame(coords_x)], ignore_index=True)
    df_coordsy = pd.concat([df_coordsy, pd.DataFrame(coords_y)], ignore_index=True)
    coords_x=pd.DataFrame({'X':coords_x})
    coords_y=pd.DataFrame({'Y':coords_y})
    df_coordsGLB_temp = coords_x.merge(coords_y, how='cross').drop_duplicates().reset_index(drop=True)
    df_coordsGLB_temp['cols']=cols_TF
    df_coordsGLB = pd.concat([df_coordsGLB, df_coordsGLB_temp], ignore_index=True).drop_duplicates().reset_index(drop=True)

# Eliminar duplicados
df_coordsx.columns=['Coords X']
df_coordsx = df_coordsx.sort_values(by='Coords X')
df_coordsx = df_coordsx.drop_duplicates().reset_index(drop=True)
df_coordsx['Esp X'] = df_coordsx['Coords X'] - df_coordsx['Coords X'].shift(1)

df_coordsy.columns=['Coords Y']
df_coordsy = df_coordsy.sort_values(by='Coords Y')
df_coordsy = df_coordsy.drop_duplicates().reset_index(drop=True)
df_coordsy['Esp Y'] = df_coordsy['Coords Y'] - df_coordsy['Coords Y'].shift(1)

#Obtener valores
Distx=df_coordsx['Coords X'].max().round(3)
Disty=df_coordsy['Coords Y'].max().round(3)
Espaciamientox=df_coordsx['Esp X'].max().round(3)
Espaciamientoy=df_coordsy['Esp Y'].max().round(3)
df_coordsGLB=df_coordsGLB[df_coordsGLB['cols']==True]
ctd_cols=df_coordsGLB.shape[0]
#------------------------------------------------------------------------------------------
#initialize etabs
SapModel=initial_etabs()
ret = SapModel.File.NewGridOnly(Num_Pisos, Altura_piso_tipico, Altura_piso1, 2, 2, 1, 1)
def_GridSystem(SapModel,df_coordsx,df_coordsy)

#define the response espectrum
nom_esp=f'Espectro_E030_Z{Z_sis}S{S_sis}R{R_sis}U{U_sis}'
define_espectrum(SapModel,nom_esp,'030')

#Configuration
e_losa=sup_cel(2*(Espaciamientox+Espaciamientoy)/180,0.1)
P_carga=Espaciamientox*Espaciamientoy*((w_m+0.4+2.4*e_losa)*(Num_Pisos-1)+w_m_azo+w_sc_azo+2.4*e_losa)             #se asume la s/c más grande: w_s/c=0.3
bl_viga=[sup_cel(max(Espaciamientox,Espaciamientoy)/12,0.1),sup_cel(sup_cel(max(Espaciamientox,Espaciamientoy)/12,0.1)*0.5,0.1)]
bl_vigac=[sup_cel(max(Espaciamientox,Espaciamientoy)/8,0.1),sup_cel(sup_cel(max(Espaciamientox,Espaciamientoy)/8,0.1)*0.5,0.1)]
l_c1=sup_cel(100*np.sqrt(P_carga/(0.35*fc_concreto_p1*9.8420653098757)),5)/100

SapModel=conf_etabs(SapModel,nom_esp)
#------------------------------------------------------------------------------------------
#add frame object by coordinates
for i in range(df_largo):
    list_planta=list(list(df.iloc[i,:]))
    planta_BF(SapModel,*list_planta)

df_losa=obtencion_db(SapModel,'Area Assignments - Floor Auto Mesh Options')
df_losa['MeshOption']='No Auto Mesh'

TableKey='Area Assignments - Floor Auto Mesh Options'
TableVersion = 1
FieldsKeysIncluded=[
'Story',
'Label',
'UniqueName',
'MeshOption',
'Restraints',
]
NumberRecords=1
TableData=[]
for i in range(0,df_losa.shape[0]):
    TableData+=list(df_losa.iloc[i,:])
ret=SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
ret=SapModel.DatabaseTables.ApplyEditedTables(True) 
#------------------------------------------------------------------------------------------
ret=SapModel.File.Save(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Base_fija')
ret=SapModel.Analyze.RunAnalysis()
ret=SapModel.View.RefreshView(0, False) # Actualizar Vista
#------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------
amor_i=0.15
coef_tb=3
[K_ten,C_disp]=KC_calc(SapModel,amor_i,coef_tb)
cte_b=1.65/(2.31-0.41*np.log(amor_i*100))
#------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------
#initialize etabs
SapModel1=initial_etabs()
#create grid
ret = SapModel1.File.NewGridOnly(Num_Pisos+2, Altura_Base, Altura_piso_tipico, 2, 2, 1, 1)
iter=1

story_name=[]
story_heights=[]
for i in range(1,Num_Pisos+1):
    story_name+=['Story'+str(i)]
    story_heights+=['Story'+str(i)]

direct_stories={'StoryNames':["Ref1","Plataforma"]+ story_name,
                'StoryHeights':[Altura_link,Altura_Base,Altura_piso1]+[Altura_piso_tipico]*(Num_Pisos-1),
                'IsMasterStory':[False,False,True]+[False]*(Num_Pisos-2)+[True],
                'SimilarToStory':["None"]*3+["Story1"]*(Num_Pisos-2)+["None"],
                'SpliceAbove':[False]*(Num_Pisos+2),
                'SpliceHeight': [0]*6,
                'Color':[255]*6
                }

ret = SapModel1.Story.SetStories_2(0, Num_Pisos+2, 
                                direct_stories['StoryNames'], direct_stories['StoryHeights'], 
                                direct_stories['IsMasterStory'], direct_stories['SimilarToStory'],
                                direct_stories['SpliceAbove'], direct_stories['SpliceHeight'], 
                                direct_stories['Color'])

def_GridSystem(SapModel1,df_coordsx,df_coordsy)

#define the link property
MyDOF=[True]*3+[False]*3
MyFixed=[True]+[False]*5
MyNonLinear=[False]*6
MyKe=[0]+[K_ten]*2+[0]*3
MyCe=[0]+[C_disp]*2+[0]*3
MyK=[amor_i]*6
MyYield=[4]*6
MyRatio=[5]*6

ret = SapModel1.PropLink.SetRubberIsolator("Aislador01", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MyYield, MyRatio, 0, 0)
ret = SapModel1.PropLink.SetWeightAndMass("Aislador01", 0.5, 0.05, 0, 0, 0)
ret = SapModel1.PropLink.SetSpringData("Aislador01", 0.001, 1)

#define the response espectrum
nom_esp=f'Espectro_E031_Z{Z_sis}S{S_sis}R{R_sis}U{U_sis}'
define_espectrum(SapModel1,nom_esp,'031')

#Configuration
SapModel1=conf_etabs(SapModel1,nom_esp)
#------------------------------------------------------------------------------------------
#add frame object by coordinates
iter=1
for i in range(0,df_largo):
    list_planta=list(list(df.iloc[i,:]))
    planta_BA(SapModel1,*list_planta)
    
df_losa1=obtencion_db(SapModel1,'Area Assignments - Floor Auto Mesh Options')
df_losa1['MeshOption']='No Auto Mesh'

TableKey='Area Assignments - Floor Auto Mesh Options'
TableVersion = 1
FieldsKeysIncluded=[
'Story',
'Label',
'UniqueName',
'MeshOption',
'Restraints',
]
NumberRecords=1
TableData=[]
for i in range(0,df_losa1.shape[0]):
    TableData+=list(df_losa1.iloc[i,:])
ret=SapModel1.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
ret=SapModel1.DatabaseTables.ApplyEditedTables(True) 
#------------------------------------------------------------------------------------------
ret=SapModel1.File.Save(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Base_aislada')
ret=SapModel1.Analyze.RunAnalysis()
ret=SapModel1.View.RefreshView(0, False) # Actualizar Vista
#------------------------------------------------------------------------------------------
#Derivas inelásticas
##Obtención de data
d_ines_max=deriva_ines_max(SapModel1)
l_c1_i=l_c1
n=1
m=1

t_control=10

while True:
    # Etabs Model of rigid base
    ret = SapModel.SetModelIsLocked(False)
    
    if abs(l_c1-bl_vigac[0])<=0.2:
        l_c1=l_c1_i
        m=1
        coef_tb=3+0.2*n
        n+=1

    direct_secc={'C1':[l_c1]*2}

    for key in direct_secc:
        def_secc(SapModel,key,nom_conc,nom_ac,direct_secc[key])

    for l in range(1,df_largo+1):
        for k in range(0, Num_Pisos + 1):
            [numx,numy,Espaciamientox,Espaciamientoy,coord_x,coord_y,vigas_TF,cols_TF,carga_sc]=list(list(df.iloc[l-1,:]))
            for i in range(1, numx + 1):
                for j in range(1, numy + 1):
                    name=str(l)+'_'+str(i)+'_'+str(j)+'_'+str(k) 
                    if vigas_TF:
                        if k==0:
                            sec_v='V2'
                            z_coord_v=0
                        else:
                            sec_v='V1'

                        if i!=1:
                            ret = SapModel.FrameObj.SetEndLengthOffset(sec_v+'x'+name, True,0,0,0)
                        if j!=1:
                            ret = SapModel.FrameObj.SetEndLengthOffset(sec_v+'y'+name, True,0,0,0)

    ret=SapModel.File.Save(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Base_fija')
    ret=SapModel.Analyze.RunAnalysis()
    ret=SapModel.View.RefreshView(0, False) # Actualizar Vista

    amor_i=0.15
    [K_ten,C_disp]=KC_calc(SapModel,amor_i,coef_tb)

    # Etabs Model of isolated base
    ret = SapModel1.SetModelIsLocked(False)

    MyDOF=[True]*3+[False]*3
    MyFixed=[True]+[False]*5
    MyNonLinear=[False]*6
    MyKe=[0]+[K_ten]*2+[0]*3
    MyCe=[0]+[C_disp]*2+[0]*3
    MyK=[amor_i]*6
    MyYield=[4]*6
    MyRatio=[5]*6

    ret = SapModel1.PropLink.SetRubberIsolator("Aislador01", MyDOF, MyFixed, MyNonLinear, MyKe, MyCe, MyK, MyYield, MyRatio, 0, 0)
    ret = SapModel1.PropLink.SetWeightAndMass("Aislador01", 0.5, 0.05, 0, 0, 0)
    ret = SapModel1.PropLink.SetSpringData("Aislador01", 0.001, 1)

    for key in direct_secc:
        def_secc(SapModel1,key,nom_conc,nom_ac,direct_secc[key])
    
    for l in range(1,df_largo+1):
        for k in range(1, Num_Pisos + 3):
            [numx,numy,Espaciamientox,Espaciamientoy,coord_x,coord_y,vigas_TF,cols_TF,carga_sc]=list(list(df.iloc[l-1,:]))
            for i in range(1, numx + 1):
                for j in range(1, numy + 1):
                    name=str(l)+'_'+str(i)+'_'+str(j)+'_'+str(k) 
                    if vigas_TF and k>2:
                        if k==0:
                            sec_v='V2'
                        else:
                            sec_v='V1'

                        if i!=1:
                            ret = SapModel1.FrameObj.SetEndLengthOffset(sec_v+'x'+name, True,0,0,0)
                        if j!=1:
                            ret = SapModel1.FrameObj.SetEndLengthOffset(sec_v+'y'+name, True,0,0,0)

    ret=SapModel1.File.Save(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Base_aislada')
    ret=SapModel1.Analyze.RunAnalysis()
    ret=SapModel1.View.RefreshView(0, False) # Actualizar Vista

    df_t=obtencion_db(SapModel,'Modal Periods And Frequencies')
    df_t=df_t[df_t['Case']=='Modal']
    periodo_1=float(df_t['Period'].iloc[0])

    df_t=obtencion_db(SapModel1,'Modal Periods And Frequencies')
    df_t=df_t[df_t['Case']=='Modal']
    periodo_2=float(df_t['Period'].iloc[0])
    
    d_ines_max=deriva_ines_max(SapModel1)
    
    print(f'Iteración de n y m:\t\t\t\t {n} y {m}')
    print(f'Dimensiones de C1 y C2 (var):\t\t\t {l_c1} y {bl_vigac[0]} ({abs(l_c1-bl_vigac[0])})')
    print(f'T_BF, T_obj (coef_tb) y T_BA (var T_obj y T_BA): {periodo_1}, {periodo_1*coef_tb} ({coef_tb}) y {periodo_2} ({periodo_2-periodo_1*coef_tb})')
    print(f'k (tonf/m) y c (tonf*s/m):\t\t\t {K_ten} y {C_disp}')
    print(f'Deriva Ines Max:\t\t\t\t {d_ines_max}')
    print('-------------------------------------------------')
    
    if (d_ines_max <= 3.5 and 0 <= abs(periodo_2 - periodo_1 * coef_tb) <= 0.15) or n >= 12:
        break
    m += 1
    l_c1+=0.1
    #----------------------------------------------
    #eliminar cuando se corrija el modelado
    if d_ines_max <= 3.5 and 0 <= abs(periodo_2 - periodo_1 * coef_tb) < t_control:
        
        t_control=abs(periodo_2 - periodo_1 * coef_tb)
        
        P_t_control=P_t
        T_obj_control=T_obj
        coef_tb_control=coef_tb,
        K_ten_control=K_ten
        C_disp_control=C_disp
        l_c1_control=l_c1
    #----------------------------------------------


#definición del análisis de cargas sísmicas estáticas
df_periods=obtencion_db(SapModel,'Modal Participating Mass Ratios')
df_periods=df_periods[df_periods['Case']=='Modal']
df_periods=df_periods.loc[:,['Mode', 'Period', 'UX', 'UY', 'UZ']]

periodo_x=df_periods[df_periods['UX']==df_periods['UX'].max()]['Period'].iloc[0]
periodo_x=float(periodo_x)

periodo_y=df_periods[df_periods['UY']==df_periods['UY'].max()]['Period'].iloc[0]
periodo_y=float(periodo_y)

[cons_z,cons_s,cons_tp,cons_tl] = const_sis(Z_sis, S_sis)
direct_userCoefficient=pd.DataFrame({
'IsAuto':['No']*2,
'XDir':['Yes','No'],
'XDirPlusE':['No']*2,
'XDirMinusE':['No']*2,
'YDir':['No','Yes'],
'YDirPlusE':['No']*2,
'YDirMinusE':['No']*2,
'EccRatio':['0.05']*2,
'TopStory':['Story'+str(Num_Pisos)]*2,
'BotStory':['Base']*2,
'OverStory':['']*2,
'OverDiaph':['']*2,
'OverEcc':['']*2,
'C':[str(calcular_C(periodo_x,cons_tp,cons_tl,'030')),str(calcular_C(periodo_y,cons_tp,cons_tl,'030'))],
'K':[str(calcular_T(periodo_x)),str(calcular_T(periodo_y))]
},index=['CSx','CSy'])
TableKey='Load Pattern Definitions - Auto Seismic - User Coefficient'
TableVersion = 1
FieldsKeysIncluded=[
'Name',
'IsAuto',
'XDir',
'XDirPlusE',
'XDirMinusE',
'YDir',
'YDirPlusE',
'YDirMinusE',
'EccRatio',
'TopStory',
'BotStory',
'OverStory',
'OverDiaph',
'OverEcc',
'C',
'K'
]
NumberRecords=1
GroupName=''
TableVersion=1  
TableData=['CSx']+direct_userCoefficient.loc['CSx',:].to_list()+['CSy']+direct_userCoefficient.loc['CSy',:].to_list()

ret = SapModel.SetModelIsLocked(False)
ret = SapModel.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion,FieldsKeysIncluded, NumberRecords,TableData)    
FillImport = True
ret= SapModel.DatabaseTables.ApplyEditedTables(FillImport) 
ret=SapModel.Analyze.RunAnalysis()

##Obtención de data
df_masa=obtencion_db(SapModel1,'Mass Summary by Story')
##Creación de la base de datos
df_masa['Masa']=df_masa['UX'].astype(float)
df_masa=df_masa.drop(['UX','UY','UZ'],axis=1)
df_masa['Masa acumulada']=df_masa['Masa'].cumsum()
P_t=df_masa['Masa acumulada'].max()*9.806

#----------------------------------------------
#eliminar cuando se corrija el modelado
P_t=P_t_control
T_obj=T_obj_control
coef_tb=coef_tb_control
K_ten=K_ten_control
C_disp=C_disp_control
l_c1=l_c1_control
#----------------------------------------------

df_results = pd.DataFrame({
    'Variables': [
        Espaciamientox, Espaciamientoy, 
        Distx, Disty, 
        ctd_cols, 
        P_t, T_obj, coef_tb,
        K_ten, C_disp,
        [l_c1]*2, 
        [bl_vigac[0] + 0.1]*2, bl_viga, bl_vigac
    ]
}, index=[
    'Espaciamiento en X', 'Espaciamiento en Y',
    'Distancia en X', 'Distancia en Y',
    'Cantidad de columnas',
    'Peso total de la estructura (tonf)', 'Periodo Objetivo (s)', 'Factor de periodo objetivo',
    'Rigidez del aislador (tonf/m)', 'Amortiguamiento del dispositivo',
    'Dimension de C1 (m)',
    'Dimension de C2 (m)', 'Dimension de V1 (m)','Dimension de V2 (m)'
])
##Hallar distancia mediana en X e Y

# Save to CSV file with tab as a separator
df_results.to_csv(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Resultados_BA.txt', sep='\t', index=True)
df_coordsx.to_csv(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Coords_X.txt', sep='\t')
df_coordsy.to_csv(r'C:\Users\migue\OneDrive\Desktop\2024-2\Python\Python_311\Etabs\Coords_Y.txt', sep='\t')