VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPropertyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'written by Andrew Reed 1996-2001
'WinHunter, a lottery prediction and statistical analysis toolkit
'Copyright (C)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful, but
'WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'
'Contact via snailmail:
'
'Andrew Reed
'7870 Almarante Place
'Laurel Hill, FL    32567
'
'email via:
'winhunter@winhunter.freeservers.com
'
'Please see the LICENSE.TXT file for the complete license

Public Event Changed()

Private mName As String         'displayed name
Private mKey As String          'Key name
Private mValue As Variant       'value to save
Private mMin As Long            'Min Value
Private mMax As Long            'max value
Private mGroup As Integer       'group this value with other values (group number)
Private mHuntable As Boolean    'determine whether or not
                                'this value should be altered by WINHunter


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property


Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(sKey As String)
    mKey = sKey
End Property


Public Property Get Value() As Variant
    Value = mValue
End Property
Public Property Let Value(vntValue As Variant)
    RaiseEvent Changed
    mValue = vntValue
End Property


Public Property Get Group() As Integer
    Group = mGroup
End Property
Public Property Let Group(intGroup As Integer)
    mGroup = intGroup
End Property

Public Property Get Min() As Long
    Min = mMin
End Property
Public Property Let Min(lMin As Long)
    mMin = lMin
End Property

Public Property Get Max() As Long
    Max = mMax
End Property
Public Property Let Max(lMax As Long)
    mMax = lMax
End Property

Public Property Get Huntable() As Boolean
    Huntable = mHuntable
End Property
Public Property Let Huntable(bPossible As Boolean)
    mHuntable = bPossible
End Property

Private Sub Class_Initialize()

    mGroup = -1

End Sub
