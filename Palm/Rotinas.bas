Attribute VB_Name = "Rotinas"
Option Explicit
 
Public Linha(99) As Variant
Public Tipo(99) As Variant
Public PI(99) As Variant
Public PF(99) As Variant
Public VM(99) As Variant
Public t(99) As Variant
Public tn(99) As Variant
Public X As Integer
Public U As Integer
Public Cont As Integer
Public HoraLargada As String
Public ArqRelat�rio As String

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1 'modo as�ncrono. toca musica sem parar
Public Const SND_LOOP = &H8

