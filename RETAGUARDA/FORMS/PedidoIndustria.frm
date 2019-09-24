VERSION 5.00
Begin VB.Form frmPedidoIndustria 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPedidoIndustria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Se uma parede tem 7 metros de comprimento, então a parede mede 7 metros lineares e não 8 metros.

'O que se compara aqui neste artigo é a diferença entre metro linear e metro quadrado.

'Imagine um rectangulo de 6 x 4 metros.

'Quantos metros lineares tem o retangulo?
'R: 6+4+6+4= 20 metros lineares – So tem de medir o comprimento

'Quantos metros quadrados tem o retangulo?

'R:Tem de calcular a área de um retangulo

'http://engiobra.com/calculadoras/areas/retangulo/

'24 metros quadrados

'========================================
'CALCULO DO METRO QUADRADO = LARGURA * ALTURA
