 If iProperties.Material = "Titanium" Then
' Get the current Part document.
Dim partDoc As PartDocument = ThisDoc.Document
' Get the TransientBRep and TransientGeometry objects.
Dim transBRep As TransientBRep = ThisApplication.TransientBRep
Dim transGeom As TransientGeometry = ThisApplication.TransientGeometry
' Combine all bodies in Part into a single transient Surface Body.
Dim combinedBodies As SurfaceBody = Nothing
For Each surfBody As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
	If combinedBodies Is Nothing Then
		combinedBodies = transBRep.Copy(surfBody)
	Else
		transBRep.DoBoolean(combinedBodies, surfBody, BooleanTypeEnum.kBooleanTypeUnion)
	End If
Next

' Get the oriented mininum range box of all bodies in Part.
' NOTE: "OrientedMinimumRangeBox" was added in Inventor 2020.3/2021.
Dim minBox As OrientedBox = combinedBodies.OrientedMinimumRangeBox

' Get length of each side of mininum range box.
Dim dir1 As Double = minBox.DirectionOne.Length
Dim dir2 As Double = minBox.DirectionTwo.Length
Dim dir3 As Double = minBox.DirectionThree.Length

' Convert lengths to document's length units.
Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure

dir1 = uom.ConvertUnits(dir1, "cm", uom.LengthUnits)
dir2 = uom.ConvertUnits(dir2, "cm", uom.LengthUnits)
dir3 = uom.ConvertUnits(dir3, "cm", uom.LengthUnits)

' Sort lengths from smallest to largest.
Dim lengths As New List(Of Double) From {dir1, dir2, dir3 }
lengths.Sort

Dim minLength As Double = lengths(0)
Dim midLength As Double = lengths(1)
Dim maxLength As Double = lengths(2)



	iProperties.Value("Custom", "H1") = Round(minLength ,0) 

	iProperties.Value("Custom", "W1") = Round(midLength ,0) 

	iProperties.Value("Custom", "L1") = Round(maxLength, 0)
	
			
				
				
				
				
				
				If Round(minLength, 0) = 4 And Round(maxLength, 0) <= 2400 And Round(midLength, 0) <= 1400 Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON SÓNG C - 3 LỚP (L:2400, W:1400, T:4)"
					iProperties.Value("Project", "Part Number") ="PA-CC-0110-08"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "11"
					
				
				End If
				
				
				
				
				
				
				If Round(minLength, 0) = 4 And Round(maxLength, 0) > 2400 Or Round(midLength, 0) > 1400 Then
					iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
					MessageBox.Show("CARTON OVER SIZE", "Thông báo")
					iProperties.Value("Project", "Part Number") ="PA-CC-0110-08"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
				
					
				
				End If
				
				
				
				
				
				
				
				If Round(minLength, 0) = 7 And Round(maxLength, 0) <= 3000 And Round(midLength, 0) <= 2180 Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON (275LBS, 48LBS/IN) - CARTON SÓNG BC - 275LBS, 48LBS/IN (L:3000, W:2180, T:7)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0110-05"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					
					iProperties.Value("Custom", "ss") = "12"
						End If
				
				
				
If Math.Round(minLength, 0) = 7 Then
    ' Kiểm tra điều kiện Round(maxLength, 0) > 3000 hoặc Round(midLength, 0) > 2180
    If Math.Round(maxLength, 0) > 3000 Or Math.Round(midLength, 0) > 2180 Then
        iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
        iProperties.Value("Project", "Part Number") = "PA-CC-0110-05"
        MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
        iProperties.Value("Custom", "MATERIAL") = "CARTON"
    End If
End If
				
				
				
				
				
				
				
				
				
				
				If Round(minLength ,0) = 10 And Round(maxLength, 0) <= 2000 And Round(midLength ,0) <= 1000 Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - Ø13 (L:2000, W:1000, T:10)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-06"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					
					iProperties.Value("Custom", "ss") = "14"
				
					
				End If
				
																If Math.Round(minLength, 0) = 10 Then
    
    															If Math.Round(maxLength, 0) > 2000 Or Math.Round(midLength, 0) > 1000 Then
       															 iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
       															 MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
      															  iProperties.Value("Project", "Part Number") = "PA-CC-0209-06"
       														 iProperties.Value("Custom", "MATERIAL") = "CARTON"
   															 End If
																Else
    
																	End If
				
				
				
				
				
				
				
				
				If Round(minLength ,0) = 15 And Round(maxLength, 0) <= 2000 And Round(midLength ,0) <= 1000 Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - Ø13 (L:2000, W:1000, T:15)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-04"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "16"
				End If
				


If Math.Round(minLength, 0) = 15 Then
    
    If Math.Round(maxLength, 0) > 2000 Or Math.Round(midLength, 0) > 1000 Then
        iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
        MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
        iProperties.Value("Project", "Part Number") = "PA-CC-0209-04"
        iProperties.Value("Custom", "MATERIAL") = "CARTON"
    End If
End If
				
				
				
				
				
				
				
				
				
				If Round(minLength ,0) = 20	And Round(maxLength, 0) <= 2400 And Round(midLength ,0) <= 1200  Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - 175LBS (L:2400, W:1200, T:20)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-02"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "18"
				End If
				
				
If Math.Round(minLength, 0) = 20 Then
    
    If Math.Round(maxLength, 0) > 2400 Or Math.Round(midLength, 0) > 1200 Then
        iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
        MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
        iProperties.Value("Project", "Part Number") = "PA-CC-0209-02"
        iProperties.Value("Custom", "MATERIAL") = "CARTON"
    End If
End If
				
				
				
				
				If Round(minLength ,0) = 30 And Round(maxLength, 0) <= 2400 And Round(midLength ,0) <= 1200 Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - Ø13 (L:2400, W:1200, T:30)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-03"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "20"
				End If
				
				
				
If Math.Round(minLength, 0) = 30 Then
   
    If Math.Round(maxLength, 0) > 2400 Or Math.Round(midLength, 0) > 1200 Then
        iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
        MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
        iProperties.Value("Project", "Part Number") = "PA-CC-0209-03"
        iProperties.Value("Custom", "MATERIAL") = "CARTON"
    End If
End If
				
				
				
				
				
				
				If Round(minLength ,0) = 50  And Round(maxLength, 0) <= 2000 And Round(midLength ,0)  <= 1000   Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - Ø13 (L:2000, W:1000, T:50)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-07"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "24"
				End If
				
				
				
				
				
If Math.Round(minLength, 0) = 50 Then
    ' Kiểm tra điều kiện Round(maxLength, 0) > 2000 hoặc Round(midLength, 0) > 1000
    If Math.Round(maxLength, 0) > 2000 Or Math.Round(midLength, 0) > 1000 Then
        iProperties.Value("Custom", "DESCRIPTION") = "QUÁ KHỔ"
        MsgBox("CARTON OVER SIZE", vbInformation, "Thông báo")
        iProperties.Value("Project", "Part Number") = "PA-CC-0209-07"
        iProperties.Value("Custom", "MATERIAL") = "CARTON"
    End If
End If
				
			
			If H1 = 30 And Round(maxLength, 0) Then
					iProperties.Value("Custom", "DESCRIPTION") = "TẤM CARTON - CARTON TỔ ONG - Ø13 (L:2400, W:1200, T:30)"
					iProperties.Value("Project", "Part Number") = "PA-CC-0209-03"
					iProperties.Value("Custom", "MATERIAL") = "CARTON"
					iProperties.Value("Custom", "ss") = "20"
					End If
			
			
End If
