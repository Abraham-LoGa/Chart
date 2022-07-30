from tkinter import *
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rc
import pandas as pd
from pandas import ExcelWriter
#from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rc

  # Creamos ventana
ventana=Tk()
ventana.title('Calculador')
ventana.geometry("530x560")

  # Funciones para texto 
def label_box(Name_object,Text,cord_x,cord_y,tipo_d_Letra,tamano):
	LabelG=Label(Name_object,text=Text)
	LabelG.place(x=cord_x,y=cord_y)
	LabelG.configure(font=(tipo_d_Letra,tamano))
def label_color(Name_object,Text,b_g,cord_x,cord_y,tipo_d_Letra,tamano):
	LabelG=Label(Name_object,text=Text,bg=b_g)
	LabelG.place(x=cord_x,y=cord_y)
	LabelG.configure(font=(tipo_d_Letra,tamano))
  # Función para Botones
def boton(Name_object,Text,llamar,cord_x,cord_y,w,h):
	botonG=Button(Name_object,text=Text,command=llamar,width=w,height=h)
	botonG.place(x=cord_x,y=cord_y)
  # Titulo
Label_1=label_box(ventana,"GRAFICADORA DE",160,20,"Arial",16)
Label_2=label_box(ventana,"DATOS PSICROMÉTRICOS",130,50,"Arial",16)
  #　Instrucciones
L_0=label_box(ventana,"----------- Instrucciones -----------",70,90,"Courier",12)
L_0=label_box(ventana,"1. El archivo debe ser un excel.",10,110,"Courier",11)
L_0=label_box(ventana,"2. Debe tener datos de 90 días.",10,130,"Courier",11)
L_0=label_box(ventana,"3. Los datos deben ser de cada diez minutos.",10,150,"Courier",11)
L_0=label_box(ventana,"4. En el archivo solo deben venir dos columnas con los",10,170,"Courier",11)
L_0=label_box(ventana,"   datos, y ser nombradas al inicio como: ",10,190,"Courier",11)
L_0=label_box(ventana,"   T_C= Temperatura en °C y P_H= Porcentaje de humedad",10,210,"Courier",11)
L_0=label_box(ventana,"5. El archivo debe encontrarse guardado en el mismo",10,230,"Courier",11)
L_0=label_box(ventana,"   lugar que el programa.",10,250,"Courier",11)
L_0=label_box(ventana,"6. Al ingresar el nombre, favor de poner la ",10,270,"Courier",11)
L_0=label_box(ventana,"   extensión '.xlsx'",10,290,"Courier",11)
L_0=label_box(ventana,"---------------------------------------------",30,320,"Courier",12)

  # Obtención de datos
L_name=label_box(ventana,"Ingrese el nombre del archivo",10,350,"Calibri",12)
name_d=Entry(ventana)
name_d.place(x=60,y=380)

L_h=label_box(ventana,"Ingrese la altura en metros",260,350,"Calibri",12)
Altura=Entry(ventana)
Altura.place(x=260,y=380)

  # Guardamos datos
def guardar():
	name=name_d.get()
	altura=float(Altura.get())
def graficar_m():
	name=name_d.get()
	z=float(Altura.get())
	ref=name
	df=pd.read_excel(ref)
	Porcentaje_H=df["P_H"]
	temp_c = df["T_C"]

	nf=df.shape[0]
	nfh=int(nf/6)
	nfd=int(nfh/24)-1
	t_C=np.zeros(nf)
	t_K=np.zeros(nf)
	p_H=np.zeros(nf)
	p_atm=101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)
	for i in range(nf):
		t_C[i]=temp_c[i]
		t_K[i]=t_C[i]+273.15
		p_H[i]=Porcentaje_H[i]

	def horas(M_t,N_F):
		N_Fh=int(N_F/6)
		M_th=np.zeros(N_Fh)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_t[i]
			if (count==5):
				aux=aux/6
				M_th[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_th
	def dias(M_h,N_F):
		N_Fd=int(N_F/24)
		M_d=np.zeros(N_Fd)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_h[i]
			if (count==23):
				aux=aux/24
				M_d[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_d
	def meses(M_d,N_F):
		N_Fm=int(N_F/29)
		M_m=np.zeros(N_Fm)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_d[i]
			if (count==28):
				aux=aux/29
				M_m[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_m
	t_Ch=horas(t_C,nf)
	t_Kh=horas(t_K,nf)
	p_Hh=horas(p_H,nf)
	t_Cd=dias(t_Ch,nfh)
	t_Kd=dias(t_Kh,nfh)
	p_Hd=dias(p_Hh,nfh)
	t_Cm=meses(t_Cd,nfd)
	t_Km=meses(t_Kd,nfd)
	p_Hm=meses(p_Hd,nfd)

	def presion_v_s(t):
		a1=-5.8002206*(pow(10,3))
		a2=1.3914993
		a3=-4.8640239*(pow(10,-2))
		a4=4.1764768*(pow(10,-5))
		a5=-1.4452093*(pow(10,-8))
		a6=6.5459673
		p_v_s=np.exp(a1/t+a2+a3*t+a4*(pow(t,2))+a5*(pow(t,3))+a6*np.log(t))
		return p_v_s
	def presion_parcial(P_V_S,P_H):
		p_p_v=(P_V_S)*(P_H)/100
		return p_p_v
	def razon_humedad(p,patm):
		r_w=0.622*(p/(patm-p))
		return r_w
	pvs_m=presion_v_s(t_Km)
	pv_m=presion_parcial(pvs_m,p_Hm)/1000
	W_m=razon_humedad(pv_m,p_atm)
	#######
	pvs_d=presion_v_s(t_Kd)
	pv_d=presion_parcial(pvs_d,p_Hd)/1000
	W_d=razon_humedad(pv_d,p_atm)
	

	Tc=np.zeros(19)
	Tc2=np.zeros(36)
	Hre=np.zeros((36,10))
	W_1=np.zeros((36,10))
	h_enta=np.zeros((36,16))
	v_es=np.zeros((36,10))
	W=np.zeros(19)
	h_v=np.zeros(36)
	v_e=np.zeros(19)
	P=np.zeros(19)
	p= 101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)*1000
	pA= 101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)*7.501
	c=0
	prc=100

	def presion_v_s(t):
		a1=-5.8002206*(pow(10,3))
		a2=1.3914993
		a3=-4.8640239*(pow(10,-2))
		a4=4.1764768*(pow(10,-5))
		a5=-1.4452093*(pow(10,-8))
		a6=6.5459673
		p_v_s=np.exp(a1/t+a2+a3*t+a4*(pow(t,2))+a5*(pow(t,3))+a6*np.log(t))
		return p_v_s

	for i in range(1,19):
		c=c+5.0
		Tc[i]=c
		k=c+273.15
		P[i]=presion_v_s(k)
		P[i]=P[i]*0.007501
	Tc[0]=0.01
	k=Tc[0]+273.15
	P[0]=presion_v_s(k)
	P[0]=P[0]*0.007501

	for i in range(10):
		for j in range(19):
			Hre[j][i]=P[j]*prc/100
		prc=prc-10

	for i in range(10):
		for j in range(19):
			W_1[j][i]=(Hre[j][i]*0.622)/(pA-Hre[j][i])
	c2=0

	  # Creación de matriz de temeperatuera para 36 datos
	for i in range(1,36):
		c2=c2+5.0
		Tc2[i]=c2
	Tc2[0]=0.01
	R_Hf=10
      # Entalpía
	for i in range(16):
		for j in range(36):
			h_enta[j][i]=((R_Hf)/4.18-(0.24*Tc2[j]))/(0.46*Tc2[j]+597.2)
		R_Hf=R_Hf+10
      # Volumen específico
	Range=0.8
	for i in range(10):
		for j in range(19):
			v_es[j][i]=18*(Range*((p)/101325)/(0.082*(Tc[j]+273.15))-1/29)
		Range=Range+0.05
	  # Gráficar
	plt.figure()
	plt.title('Volúmen Específico m^3/kg (0.8-1.25)', fontsize=10)
	plt.suptitle('Carta Psicrométrica')
	ax = plt.gca()
	ax2 = plt.gca()
	  # Configuración y nombramiento para la gráfica
	ax.grid(True)
	ax.set_xticks(np.linspace(0,90,19))
	ax.set_xlim(0,50)
	ax.set_xlabel("Temperatura de Bulbo Seco (°C)")
	ax.set_yticks(np.linspace(0,130,14))
	ax.set_ylim(0,130)
	ax.set_ylabel("Entalpía (kJ/kg)")
	ax2 = ax.twinx()
	ax2.set_ylabel(r"Relación Humedad Kg de agua/Kg de aire seco" )
	ax2.set_ylim(0,0.05,11)
	  # Graficar datos
	for i in range (10):
		for j in range (19):
			v_e[j]=v_es[j][i]
		ax2.plot(Tc,v_e,"g--",linewidth=1)
	for i in range (16):
		for j in range (36):
			h_v[j]=h_enta[j][i]
		ax2.plot(Tc2,h_v,"r--",linewidth=1)
	for i in range (10):
		for j in range (19):
			W[j]=W_1[j][i]
		ax2.plot(Tc,W)

	ax2.plot(t_Cd[0],W_m[0],"bo",markersize=7)
	ax2.plot(t_Cd[1],W_m[1],"ro",markersize=7)
	ax2.plot(t_Cd[2],W_m[2],"go",markersize=7)
	plt.show()

def graficar_d():
	name=name_d.get()
	z=float(Altura.get())
	ref=name
	df=pd.read_excel(ref)
	Porcentaje_H=df["P_H"]
	temp_c = df["T_C"]

	nf=df.shape[0]
	nfh=int(nf/6)
	nfd=int(nfh/24)-1
	t_C=np.zeros(nf)
	t_K=np.zeros(nf)
	p_H=np.zeros(nf)
	p_atm=101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)
	for i in range(nf):
		t_C[i]=temp_c[i]
		t_K[i]=t_C[i]+273.15
		p_H[i]=Porcentaje_H[i]

	def horas(M_t,N_F):
		N_Fh=int(N_F/6)
		M_th=np.zeros(N_Fh)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_t[i]
			if (count==5):
				aux=aux/6
				M_th[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_th
	def dias(M_h,N_F):
		N_Fd=int(N_F/24)
		M_d=np.zeros(N_Fd)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_h[i]
			if (count==23):
				aux=aux/24
				M_d[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_d
	def meses(M_d,N_F):
		N_Fm=int(N_F/29)
		M_m=np.zeros(N_Fm)
		count=0
		aux=0
		n=0
		for i in range(N_F):
			aux=aux+M_d[i]
			if (count==28):
				aux=aux/29
				M_m[n]=aux
				n=n+1
				aux=0
				count=-1
			count=count+1
		return M_m
	t_Ch=horas(t_C,nf)
	t_Kh=horas(t_K,nf)
	p_Hh=horas(p_H,nf)
	t_Cd=dias(t_Ch,nfh)
	t_Kd=dias(t_Kh,nfh)
	p_Hd=dias(p_Hh,nfh)
	t_Cm=meses(t_Cd,nfd)
	t_Km=meses(t_Kd,nfd)
	p_Hm=meses(p_Hd,nfd)
	#print(t_Cm)
	#print(p_Hm)
	def presion_v_s(t):
		a1=-5.8002206*(pow(10,3))
		a2=1.3914993
		a3=-4.8640239*(pow(10,-2))
		a4=4.1764768*(pow(10,-5))
		a5=-1.4452093*(pow(10,-8))
		a6=6.5459673
		p_v_s=np.exp(a1/t+a2+a3*t+a4*(pow(t,2))+a5*(pow(t,3))+a6*np.log(t))
		return p_v_s
	def presion_parcial(P_V_S,P_H):
		p_p_v=(P_V_S)*(P_H)/100
		return p_p_v
	def razon_humedad(p,patm):
		r_w=0.622*(p/(patm-p))
		return r_w
	pvs_m=presion_v_s(t_Km)
	pv_m=presion_parcial(pvs_m,p_Hm)/1000
	W_m=razon_humedad(pv_m,p_atm)
	#######
	pvs_d=presion_v_s(t_Kd)
	pv_d=presion_parcial(pvs_d,p_Hd)/1000
	W_d=razon_humedad(pv_d,p_atm)
	
	Tc=np.zeros(19)
	Tc2=np.zeros(36)
	Hre=np.zeros((36,10))
	W_1=np.zeros((36,10))
	h_enta=np.zeros((36,16))
	v_es=np.zeros((36,10))
	W=np.zeros(19)
	h_v=np.zeros(36)
	v_e=np.zeros(19)
	P=np.zeros(19)
	p= 101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)*1000
	pA= 101.325*pow(1-2.25577*(pow(10,-5)*(z)),5.2559)*7.501
	c=0
	prc=100
	def presion_v_s(t):
		a1=-5.8002206*(pow(10,3))
		a2=1.3914993
		a3=-4.8640239*(pow(10,-2))
		a4=4.1764768*(pow(10,-5))
		a5=-1.4452093*(pow(10,-8))
		a6=6.5459673
		p_v_s=np.exp(a1/t+a2+a3*t+a4*(pow(t,2))+a5*(pow(t,3))+a6*np.log(t))
		return p_v_s

	for i in range(1,19):
		c=c+5.0
		Tc[i]=c
		k=c+273.15
		P[i]=presion_v_s(k)
		P[i]=P[i]*0.007501
	Tc[0]=0.01
	k=Tc[0]+273.15
	P[0]=presion_v_s(k)
	P[0]=P[0]*0.007501

	for i in range(10):
		for j in range(19):
			Hre[j][i]=P[j]*prc/100
		prc=prc-10

	for i in range(10):
		for j in range(19):
			W_1[j][i]=(Hre[j][i]*0.622)/(pA-Hre[j][i])
	c2=0

	  # Creación de matriz de temeperatuera para 36 datos
	for i in range(1,36):
		c2=c2+5.0
		Tc2[i]=c2
	Tc2[0]=0.01
	R_Hf=10
      # Entalpía
	for i in range(16):
		for j in range(36):
			h_enta[j][i]=((R_Hf)/4.18-(0.24*Tc2[j]))/(0.46*Tc2[j]+597.2)
		R_Hf=R_Hf+10
      # Volumen específico
	Range=0.8
	for i in range(10):
		for j in range(19):
			v_es[j][i]=18*(Range*((p)/101325)/(0.082*(Tc[j]+273.15))-1/29)
		Range=Range+0.05
	  # Gráficar
	plt.figure()
	plt.title('Volúmen Específico m^3/kg (0.8-1.25)', fontsize=10)
	plt.suptitle('Carta Psicrométrica')
	ax = plt.gca()
	ax2 = plt.gca()
	  # Configuración y nombramiento para la gráfica
	ax.grid(True)
	ax.set_xticks(np.linspace(0,90,19))
	ax.set_xlim(0,50)
	ax.set_xlabel("Temperatura de Bulbo Seco (°C)")
	ax.set_yticks(np.linspace(0,130,14))
	ax.set_ylim(0,130)
	ax.set_ylabel("Entalpía (kJ/kg)")
	ax2 = ax.twinx()
	ax2.set_ylabel(r"Relación Humedad Kg de agua/Kg de aire seco" )
	ax2.set_ylim(0,0.05,11)
	  # Graficar datos
	for i in range (10):
		for j in range (19):
			v_e[j]=v_es[j][i]
		ax2.plot(Tc,v_e,"g--",linewidth=1)
	for i in range (16):
		for j in range (36):
			h_v[j]=h_enta[j][i]
		ax2.plot(Tc2,h_v,"r--",linewidth=1)
	for i in range (10):
		for j in range (19):
			W[j]=W_1[j][i]
		ax2.plot(Tc,W)
	for k in range(28):
		ax2.plot(t_Cd[k],W_d[k],"b*",markersize=3)
	for k in range(29,57):
		ax2.plot(t_Cd[k],W_d[k],"r*",markersize=3)
	for k in range(58,87):
		ax2.plot(t_Cd[k],W_d[k],"g*",markersize=3)

	plt.show()
  # Botones para graficar
  
B1=boton(ventana,"Guardar",guardar,430,375,8,1)
L_m1=label_color(ventana," ","blue",50,430,"Calibri",5)
L_mL=label_box(ventana,"Mes 1",70,425,"Courier",12)
L_m2=label_color(ventana," ","red",210,430,"Calibri",5)
L_m2=label_box(ventana,"Mes 2",230,425,"Courier",12)
L_m3=label_color(ventana," ","green",370,430,"Calibri",5)
L_m3=label_box(ventana,"Mes 3",390,425,"Courier",12)
B2=boton(ventana,"Graficar Meses",graficar_m,130,490,11,1)
B3=boton(ventana,"Graficar Días",graficar_d,260,490,11,1)
  # Muestra ventana
ventana.mainloop()