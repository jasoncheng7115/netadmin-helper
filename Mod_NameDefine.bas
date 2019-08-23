Attribute VB_Name = "Mod_NameDefine"

'測試電腦有無回應時的限制
Public Const Timeout_Ping As Single = 3
Public Port_Ping As String

'可連線電腦總數
Public Total_Connected_Computers As Long

Const tmp_Format As String = "###,###,###,###,###,###"

Function Format_KB_By_B(Src_String) As String
    If IsNull(Src_String) = True Then
        Format_KB_By_B = "0 KB"
    Else
        Format_KB_By_B = Format((Val(Src_String) / 1024), tmp_Format) & " KB"
    End If
End Function

Function Format_MB_By_B(Src_String) As String
    If IsNull(Src_String) = True Then
        Format_MB_By_B = "0 MB"
    Else
        Format_MB_By_B = Format((Val(Src_String) / 1024 / 1024), tmp_Format) & " MB"
    End If
End Function

Function Format_MB_By_K(Src_String) As String
    
    If IsNull(Src_String) = True Then
        Format_MB_By_K = "0 MB"
    Else
        Format_MB_By_K = Format((Val(Src_String) / 1024), tmp_Format) & " MB"
    End If
    
End Function


Public Function AutoSelStr(Src_Obj As Object)
'自動選取輸入格字串

On Error Resume Next

Src_Obj.SelStart = 0
Src_Obj.SelLength = Len(Src_Obj.Text)

End Function


Public Function Change_GMT(Src_String As String) As String
'轉換含 GMT 的時間
    
    '先取出左邊日期時間
    Dim tmp_1: tmp_1 = Left(Src_String, 15)
    
    '轉換成正確時間與日期格式
    Dim tmp_Datetime As Date
    tmp_Datetime = Format(Format(tmp_1, "####/##/## ##:##:##"), "yyyy/mm/dd hh:mm:ss")
    
    
    '先取出右邊 GMT
    Dim tmp_2: tmp_2 = Right(Src_String, 10)
    
    '取得 GMT 時差
    Dim tmp_Gmt_hr
    tmp_Gmt_hr = CLng(Mid(tmp_2, InStr(1, tmp_2, "+") + 1))
    
    Change_GMT = Format(DateAdd("n", tmp_Gmt_hr, tmp_Datetime), "yyyy/mm/dd hh:mm:ss")
    

    
End Function


Public Function strOsLang(intOsLang As Integer) As String
'轉換語系
    
    Select Case intOsLang
       Case 1: strOsLang = " Arabic "
       Case 4: strOsLang = " Chinese "
       Case 9: strOsLang = " English "
       Case 1025: strOsLang = " Arabic (Saudi Arabia) "
       Case 1026: strOsLang = " Bulgarian "
       Case 1027: strOsLang = " Catalan "
       Case 1028: strOsLang = " Chinese (Taiwan) "
       Case 1029: strOsLang = " Czech "
       Case 1030: strOsLang = " Danish "
       Case 1031: strOsLang = " German (Germany) "
       Case 1032: strOsLang = " Greek "
       Case 1033: strOsLang = " English (United States) "
       Case 1034: strOsLang = " Spanish (Traditional Sort) "
       Case 1035: strOsLang = " Finnish "
       Case 1036: strOsLang = " French (France) "
       Case 1037: strOsLang = " Hebrew "
       Case 1038: strOsLang = " Hungarian "
       Case 1039: strOsLang = " Icelandic "
       Case 1040: strOsLang = " Italian (Italy) "
       Case 1041: strOsLang = " Japanese "
       Case 1042: strOsLang = " Korean "
       Case 1043: strOsLang = " Dutch (Netherlands) "
       Case 1044: strOsLang = " Norwegian (Bokmal) "
       Case 1045: strOsLang = " Polish "
       Case 1046: strOsLang = " Portuguese (Brazil) "
       Case 1047: strOsLang = " Rhaeto-Romanic "
       Case 1048: strOsLang = " Romanian "
       Case 1049: strOsLang = " Russian "
       Case 1050: strOsLang = " Croatian "
       Case 1051: strOsLang = " Slovak "
       Case 1052: strOsLang = " Albanian "
       Case 1053: strOsLang = " Swedish "
       Case 1054: strOsLang = " Thai "
       Case 1055: strOsLang = " Turkish "
       Case 1056: strOsLang = " Urdu "
       Case 1057: strOsLang = " Indonesian "
       Case 1058: strOsLang = " Ukrainian "
       Case 1059: strOsLang = " Belarusian "
       Case 1060: strOsLang = " Slovenian "
       Case 1061: strOsLang = " Estonian "
       Case 1062: strOsLang = " Latvian "
       Case 1063: strOsLang = " Lithuanian "
       Case 1065: strOsLang = " Farsi "
       Case 1066: strOsLang = " Vietnamese "
       Case 1069: strOsLang = " Basque "
       Case 1070: strOsLang = " Sorbian "
       Case 1071: strOsLang = " Macedonian (FYROM) "
       Case 1072: strOsLang = " Sutu "
       Case 1073: strOsLang = " Tsonga "
       Case 1074: strOsLang = " Tswana "
       Case 1076: strOsLang = " Xhosa "
       Case 1077: strOsLang = " Zulu "
       Case 1078: strOsLang = " Afrikaans "
       Case 1080: strOsLang = " Faeroese "
       Case 1081: strOsLang = " Hindi "
       Case 1082: strOsLang = " Maltese "
       Case 1084: strOsLang = " Gaelic "
       Case 1085: strOsLang = " Yiddish "
       Case 1086: strOsLang = " Malay (Malaysia) "
       Case 2049: strOsLang = " Arabic (Iraq) "
       Case 2052: strOsLang = " Chinese (PRC) "
       Case 2055: strOsLang = " German (Switzerland) "
       Case 2057: strOsLang = " English (United Kingdom) "
       Case 2058: strOsLang = " Spanish (Mexico) "
       Case 2060: strOsLang = " French (Belgium) "
       Case 2064: strOsLang = " Italian (Switzerland) "
       Case 2067: strOsLang = " Dutch (Belgium) "
       Case 2068: strOsLang = " Norwegian (Nynorsk) "
       Case 2070: strOsLang = " Portuguese (Portugal) "
       Case 2072: strOsLang = " Romanian (Moldova) "
       Case 2073: strOsLang = " Russian (Moldova) "
       Case 2074: strOsLang = " Serbian (Latin) "
       Case 2077: strOsLang = " Swedish (Finland) "
       Case 3073: strOsLang = " Arabic (Egypt) "
       Case 3076: strOsLang = " Chinese (Hong Kong SAR) "
       Case 3079: strOsLang = " German (Austria) "
       Case 3081: strOsLang = " English (Australia) "
       Case 3082: strOsLang = " Spanish (International Sort) "
       Case 3084: strOsLang = " French (Canada) "
       Case 3098: strOsLang = " Serbian (Cyrillic) "
       Case 4097: strOsLang = " Arabic (Libya) "
       Case 4100: strOsLang = " Chinese (Singapore) "
       Case 4103: strOsLang = " German (Luxembourg) "
       Case 4105: strOsLang = " English (Canada) "
       Case 4106: strOsLang = " Spanish (Guatemala) "
       Case 4108: strOsLang = " French (Switzerland) "
       Case 5121: strOsLang = " Arabic (Algeria) "
       Case 5127: strOsLang = " German (Liechtenstein) "
       Case 5129: strOsLang = " English (New Zealand) "
       Case 5130: strOsLang = " Spanish (Costa Rica) "
       Case 5132: strOsLang = " French (Luxembourg) "
       Case 6145: strOsLang = " Arabic (Morocco) "
       Case 6153: strOsLang = " English (Ireland) "
       Case 6154: strOsLang = " Spanish (Panama) "
       Case 7169: strOsLang = " Arabic (Tunisia) "
       Case 7177: strOsLang = " English (South Africa) "
       Case 7178: strOsLang = " Spanish (Dominican Republic) "
       Case 8193: strOsLang = " Arabic (Oman) "
       Case 8201: strOsLang = " English (Jamaica) "
       Case 8202: strOsLang = " Spanish (Venezuela) "
       Case 9217: strOsLang = " Arabic (Yemen) "
       Case 9226: strOsLang = " Spanish (Colombia) "
       Case 10241: strOsLang = " Arabic (Syria) "
       Case 10249: strOsLang = " English (Belize) "
       Case 10250: strOsLang = " Spanish (Peru) "
       Case 11265: strOsLang = " Arabic (Jordan) "
       Case 11273: strOsLang = " English (Trinidad) "
       Case 11274: strOsLang = " Spanish (Argentina) "
       Case 12289: strOsLang = " Arabic (Lebanon) "
       Case 12298: strOsLang = " Spanish (Ecuador) "
       Case 13313: strOsLang = " Arabic (Kuwait) "
       Case 13322: strOsLang = " Spanish (Chile) "
       Case 14337: strOsLang = " Arabic (U.A.E.) "
       Case 14346: strOsLang = " Spanish (Uruguay) "
       Case 15361: strOsLang = " Arabic (Bahrain) "
       Case 15370: strOsLang = " Spanish (Paraguay) "
       Case 16385: strOsLang = " Arabic (Qatar) "
       Case 16394: strOsLang = " Spanish (Bolivia) "
       Case 17418: strOsLang = " Spanish (El Salvador) "
       Case 18442: strOsLang = " Spanish (Honduras) "
       Case 19466: strOsLang = " Spanish (Nicaragua) "
       Case 20490: strOsLang = " Spanish (Puerto Rico) "
    
    End Select
    
    strOsLang = Trim(strOsLang)

End Function

Public Function strLocale(intLocale As String) As String
'轉換國別
    
    strLocale = Trim(strOsLang(HexToDec(intLocale)))

End Function

Public Function strCodepage(intCodepage As String) As String
'轉換字碼頁

    Select Case intCodepage
        Case 936: strCodepage = "簡體中文 GBK"
        Case 950: strCodepage = "繁体中文 BIG5"
        Case 437: strCodepage = "美國/加拿大英語"
        Case 932: strCodepage = "日文"
        Case 949: strCodepage = "韓文"
        Case 866: strCodepage = "俄文"
        Case 65001: strCodepage = "Unicode UFT-8"
    End Select
    
    strCodepage = Trim(strCodepage)
End Function

Public Function HexToDec(sByte)
  For Counter = 1 To Len(sByte)
    Select Case Mid(sByte, Counter, 1)
      Case "0":   HexToDec = HexToDec + 0 * (16 ^ (Len(sByte) - Counter))
      Case "1":   HexToDec = HexToDec + 1 * (16 ^ (Len(sByte) - Counter))
      Case "2":   HexToDec = HexToDec + 2 * (16 ^ (Len(sByte) - Counter))
      Case "3":   HexToDec = HexToDec + 3 * (16 ^ (Len(sByte) - Counter))
      Case "4":   HexToDec = HexToDec + 4 * (16 ^ (Len(sByte) - Counter))
      Case "5":   HexToDec = HexToDec + 5 * (16 ^ (Len(sByte) - Counter))
      Case "6":   HexToDec = HexToDec + 6 * (16 ^ (Len(sByte) - Counter))
      Case "7":   HexToDec = HexToDec + 7 * (16 ^ (Len(sByte) - Counter))
      Case "8":   HexToDec = HexToDec + 8 * (16 ^ (Len(sByte) - Counter))
      Case "9":   HexToDec = HexToDec + 9 * (16 ^ (Len(sByte) - Counter))
      Case "A":   HexToDec = HexToDec + 10 * (16 ^ (Len(sByte) - Counter))
      Case "B":   HexToDec = HexToDec + 11 * (16 ^ (Len(sByte) - Counter))
      Case "C":   HexToDec = HexToDec + 12 * (16 ^ (Len(sByte) - Counter))
      Case "D":   HexToDec = HexToDec + 13 * (16 ^ (Len(sByte) - Counter))
      Case "E":   HexToDec = HexToDec + 14 * (16 ^ (Len(sByte) - Counter))
      Case "F":   HexToDec = HexToDec + 15 * (16 ^ (Len(sByte) - Counter))
    End Select
  Next
End Function

Public Function StrNullToSpace(Src_String As Variant) As String
'將 Null 轉為空白

If IsNull(Src_String) Then
    StrNullToSpace = ""
Else
    StrNullToSpace = Src_String
End If

End Function

