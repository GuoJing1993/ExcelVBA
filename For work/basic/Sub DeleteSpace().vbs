{\rtf1\ansi\ansicpg950\cocoartf1345\cocoasubrtf380
{\fonttbl\f0\fswiss\fcharset0 Helvetica;\f1\fnil\fcharset136 STHeitiTC-Light;}
{\colortbl;\red255\green255\blue255;}
\paperw11900\paperh16840\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural

\f0\fs24 \cf0 Sub 
\f1 deletespace()
\f0 \
    With ActiveSheet\
        ROW1 = InputBox(\'93Row Number\'94)\
        COLUMN1 = InputBox(\'93Column Number\'94)\
        For i = ROW1 To 1 Step -1\
            For j = 1 To COLUMN1\
                If Len(.Cells(i, j)) = 0 Then\
                    .Cells(i, j).EntireRow.Delete\
                    Exit For\
                End If\
            Next j\
        Next i\
    End With\
End Sub}