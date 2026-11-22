' roMail - trimitere email simpla (VB.NET, WinForms, fara biblioteci externe)
' Compilare: build_vb.bat  (foloseste vbc.exe inclus in Windows)
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Windows.Forms

Module Program
    <STAThread()> Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Module

' ====================== iconite (PNG incorporate) ======================
Public Module Icoane
    Private ReadOnly cache As New Dictionary(Of String, Image)()
    Private ReadOnly date_ As Dictionary(Of String, String) = Construieste()

    Private Function Construieste() As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)()
        d("plus") = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABNklEQVR4nK2SMVLDQAxF31/vpHIVmAAFBVV8C5/Ct8gFmOEY0NOn4Ay+hV2loAA8KVMxuysK7JA4MQ38ZlejL83Xl8QEyrL0u91OAHmeW13XYYr7J+hMbABFUawkXQKY2bZpmscxB8BPNBDwkGXZFUAI4QN4OshNNkj9a2a2DSFcDAoOitJhwZQCJHlJ/gzvSIHKsvR1XYeiKFbOuVWMMUjyZnYnadYr+JS0MbOQZZmPMb60bXtfVVW2XxVw7ZxbppSQjr2VNJO0BHDOkVK6Bei6To4/wud5PszznlJqzSwAJyMAGzMLKSVvZq8Ai8XCfruDZpBtZm3TNMU5E8db2CfMLALh4H/CARh7MMQC5r4HMOfnao9qphQY8BxjvOnjt1Hu/zA2cY+qqrKu6wTfbq/X63iO9wXXc5gE8NaR/QAAAABJRU5ErkJggg=="
        d("check") = "iVBORw0KGgoAAAANSUhEUgAAABYAAAAWCAYAAADEtGw7AAAC/0lEQVR4nLWUvUszWRTGf/fO56Ip01iIiImV8IKFhbhTCrvCsoJuoZWVYKn5FwQLEUSwsBFlF9EqSMqVnQVZsBAsLGIgCjYiRgU1TiZz526xmWzE5PV1YQ8M3Dv33uc855znHPifTLT5Jz3Pk58F8n1fA+ozjj5j4t2isdaZTOaL67o/GoZhfOQojmMADfD6+vpHqVT6M8FJHkpADwwMfLFt+69areYEQfBNFKWUSClJpVLUarWJUqlUmJqaMkwAz/Ok7/uR4zg/BUHgjI6OBiMjI2Ycx0jZPt0NtkRRxM3NTf3g4OA713V/AQq3t7fCbL1sWZZ+eHjQnueZc3NzZjvAVtvf3yefzzM2NobWGiFEnJy9eyyEENVqlSiKiKII0zRbz4iiCMdxODo6IpfLEYYhJycndHd3E4bhvynqlDfTNDEMA9M0m58QAsdxuLq6IpfLoZRieHiYxcVFnp+f35DoqNc4jjFNs5nLRqg8PT0xPz/P3d0druuytrbG4OAgQRAghPg6cFLpx8dHpJQopYjjGCEES0tLnJ+fI4RgeXmZTCZDpVLhH3XSGVhrjWEY7O3tMT4+TrlcxjAMDMNgZWWFQqGA1pqFhQWmp6eJ4xjLst6Taxf+/f09u7u7XFxcMDs7y/X1NYeHh2xsbCCEYGJiglwuR71e7yjHN6qQUlKv10mn02xtbTE5OUm5XGZmZoZqtYrWmqGhIVZXV9Fad6hOh1RIKXl5eaGnp4ednR36+/u5vLykUqmQTqfZ3Nykq6sLrXVHtm2BW5lns1m2t7fp7e0llUqxvr5OX18fYRiitUYp1SzsV1MBaK01rus2C5LNZjk+Pqa1vW3bBmgqIYmAxkB6B6yUMh3H0WdnZ1E+n0cphRACy7IQQhCG4ZvwE2enp6eRbdtSKdWUR6JoA1CZTOZ727b9Wq3WFLzWulmo1gZI9lprLMtKWnqmWCz+5nme2W4e/+A4zs+NaD4c/MngCcPw92Kx+GuC8+7eR0Dfau2ADM/z/pMD3/ejZP03rr5PbJWV+BkAAAAASUVORK5CYII="
        d("redx") = "iVBORw0KGgoAAAANSUhEUgAAABIAAAASCAYAAABWzo5XAAACr0lEQVR4nJ2TT2sUSxTFf7erZ3TaSceJODFP+3O4MYJPEATdjEHQjQTErQu/ikzMymyMWbgIZKNbFw9diSDqxkbMaBxIMDIzHftP1X2LnslkjILvHSgoirqnTp1zr6iqLC0tHe92u8L/QBiGRFH0XdrtdgP4p1arnbDWKvBHhCKiIiJpmuaq+rff7XYliqJTrVbruDEGVQXnytueN1mtWq7huYiwvr5OHMe+P7ySGWO0UqmoKwrx6nUU0CQZkzmHHDmCGINLU0RVERERya21OnpSVFVsUYgEgeyurUnv6VORIBCXZeLyXCQIZO/VK9m6d0/U88SBqHMyssLfl+0cpl5nd22Nrdu38RoNzOoqwfw8AOnbt3xZXOTHmze4Xo+5dhvyfL98wgS1Fm96GtNooP0+nZs3SV68IPvwgc1r18g3N6nMzhKcO3fI/LEiz0PTlKnLlzGrq3Ru3MB9+8aXxUWkWqXodJBqlVP37zO9sIAdDBBjfq0Iz8P2egTz85x58gQ/irA7O9ivXzEzM/y1skK4sEDR6x1K9Kd8xzH7J08i1SqIlEXG4M/OjtvgJ0wSOYeZmiJ9945PV66QxzHe0aNIrYbd2WGz1SJ5/hwThuNeO0SkilSrJC9f0rl+nWJra9+TM48f44Uhrt/n861b9J89K9UeUDZJ5Pv8eP2a9P17pFJhbnmZ8OpVamfPcvrRI0yjQfbxI7srK+WXD2AiNZskzNy5g+7tYRoNpkfGqlI/f5655WV2Hz5k7sED1NqJofQP7NUzRl2ScOLuXRTKiIdTVAwGHLt4kWOXLqFZhlirB1WNiDwREUDE88oZg4k+EWMgy8o9MJwzVMsx88MwJMuybGNjI7fW6pDw91Ad+aMiIoPBIG82m+pHUfR9e3v7QhzHJhu++F/QbDa1KIr4X3rcMv0Baaz6AAAAAElFTkSuQmCC"
        d("trash") = "iVBORw0KGgoAAAANSUhEUgAAABYAAAAWCAYAAADEtGw7AAAC+0lEQVR4nLVUPW8kRRB9VdOz4izf7u0Z2LPEyQGWdmSJyMGdSCYjJnFICAkE/AAkJwiBEAkJf4DEKRI/oDMnzu60WB50OoTl/ThjgVbA7PTUI3CPb7y2zwTwkq4ebb969baqBLcgz3M3n89l+fvBwUEAwNve/+e4pCTPc9e+z2az18zsExFZISkAYGYQkTpN0+/W1tZO5/O5rK6uEgC89zViFVdKbGN3d1f39vbqJElQ1/W5EhGoKhaLxVtHR0fHtyre2trqhBDeT5KkQzKJJCskvwRgAOYAQPIOgK6qfkZyQlJUta7rmmb2Q1EUfwAQF8lpZnedc3tpmiKEcJHZzJrwXqMYAFT1qwt1InDOoaqqdwA8aYgBAJ1O5+/FYvFtWZYA8B5JjY/cuVCiRSwhhFpEDIADUJD8RUROIx2veJxl2YcAvlbV3k3+tUEymNnB4eHh45a9l4g1z3Mdj8cjVd0k+QXJEFWxedCc8X+giHxKMvT7/Yf7+/t/XZc8AYDhcPhjlmUVAP03irMsOx4Oh6OWWgDn/gAA8jwX7z0AzJIkcVmWbQ0Gg58mk8k3JB2AExF5aGbPVPXRYDDYOTs7u1+W5QDAz20bcJ0qVR2LCEh2vfeB5J3YYp3mBLDqvQ9lWa4kSaIiMgGAnZ2dC8VXiElORQQi0gcAETERCSJSNyeAGgDMrKeqQnIKANPp9GZiM5vGBG+0ypM40u0Yqno/ihhfqbwJvPeMCmeReLD842uqa5JPbiRu4Jz7jSRE5M1XcDaKX4/3F68ibnZrQ/wgbjsnIi5OYHNP8jx3JB8AgIi8AF5WDbTarSGuqup3VTWS97z3YWNj4+M0TVkURbW5uZkWRVFtb2/f9d6HLMt6cdRP2xwXJbVirq+vr3S73YmI/Ery89gNhpcTpwBgZk5EPlLVdwG8PRqNnkcHbFkxAMjJycmfvV7vaZqmjwB832yzZTRLabFYHPf7/Slaw3GdFRI9+8DMHtd1HZrWWoaIMEkSp6pP4o64RPy/4R99MW9pZ+MLSAAAAABJRU5ErkJggg=="
        d("env") = "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"
        d("floppy") = "iVBORw0KGgoAAAANSUhEUgAAACQAAAAkCAYAAADhAJiYAAADF0lEQVR4nM2Yu07kMBSG/xPnMoFMEAPUPAGaKeiQKEGIB6CkQzwAFT0SINFRUCNBxQPQ0oCggwaEoOLSTTOKEJjYZ4tVImZy2TjLLvNLbk4c58uxzyWh4+NjxhDJ+mmAQQ0dkF10gYgq2aqImaG1rnR/IZCUMmNTStUCsiwLjUYDWmtzICKClBJnZ2eQUqZvRUTwfR9EBObqcUBEiKII09PT6HQ6fWtWAkokpYSUEpZlIY5jBEGAjY0NBEGAOI6Ntm9nZwc3NzcgIrTb7VKo0jP0dQghEIYhms2mMZAQAp7n4fb2Fsxc6qlCoEExM+I4TocJEDODmVMoAOh0Ovj8/KwPBGS9ZqpBqJmZmcxB/+95iJnhui7u7+/R6/UghOgLkh9LjLadvzk/BlSUOoaudAwd0B+jjJkhhMD7+zu2t7dhWZZRpgaAbrcL27Yr3Vc57LXWeHl5MQJJH2LbldOEUR5yXbcWkIlHjYBMt6qOjIASmWbpf+YhAJV6mkFZVvVgNq5lSU9UpsQjybyPj48++18DERGUUvB9H6urq/B9H0qpQrDEI0mVPzo6wvPzM0ZHR/8IVstDIyMjpUBSSmit4boumBmtVgtzc3M4OTmB53mlXWetM6SUygWyLAtvb284PDxEt9vF7OwslpaWEEURVlZWMDExgf39fTSbzWLvmgIB2b7oa9/tOE5mBEGAXq+HxcVFrK+vI4qidP6gaoV9GWi73cbu7m5achzHwebmJhqNBpRSWF5eBgAcHBzkfhp9KxAAOI4Dz/P6bL7vA/jdWydQnufh4eEhc5a+HSiJrERFZ2V8fLxvqxP9k/ajqO/WWkMIgYuLC2xtbaVzv77At3koyVV3d3d9IEopnJ6eYm1tDWNjYzg/P8fe3h5c183007lAzAzbtjE/P5+ZTER4fX0tzdSDpUVrjcnJSYRhiKenJ1xfX2NhYQFCiNwEm+shIkIYhhk7M+d+85dJa42pqSk8Pj7i6uoKrVYrbfLy/hUUblkcx7l200ovhEAURbi8vAQRpYm1aC2j3zF1xcx9Fb9s7aFr8n8B7UNjDRyDwyMAAAAASUVORK5CYII="
        d("stop") = "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"
        d("app") = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFaElEQVR4nO2WW2xcVxWGv73PZWY8Mx5P7djCY8dOG3CxRZW6Ump4KG1BERUUqUFTBJSKiqJKVZVKpRHqAzIj5aEUVeImEIQKoiABNhIKokpVZLVORVKaQhq1TZTEqmPHTh0cxxmPPTPnthcPjm/JpHVJETzkl/bDOWef839nrb3X2nBd1/X/pHx+wEJEwX9v5PMD1v/6P9dIrb647au/ucsY90SopZLJZNiY+XBMikUYLxZxBStS1S0fDdIvDQ7eHwHY9PdrCgWz5WvPdYMe0lYwH1MExfPn1N8nA7RSKPV+FldXZISYa5OIOYIS5dip7Iha+CwwlM8PWDaF7wkUwE9P45YvRmJlvapPb087t3S1UlrwUIoPDCECIkI6GWdkbJqDR0axtMahIpayzgEMdr8t9tILHa1OcPJMKHf2bZJz00V549i42vqJHF/a1ovnRYRRhNZ6XeZGBK0UyYTNgddHGXh+jFxzvXRvblFDB0+oTDzpL81dBkgmHBWEocq1NKhHvtzHj/YcUN//5RBHjk3yzXwfibhLeb6CbS1CSA1jBUSRIRazEYHdg4f44wtH+eSWTnY+fDdvHJ9U+4eP4TRay/G0V39Aa0XVC3AszY4H7+DmG5v51cAhxs/OsuPBO7ipo4nSvIdSCl0jJZExZNIJps7P8dPfvsLR45N8476tbN92C3UJh3LFR12WyytiqrUiMsJC2efzd/aw64kvUPUDnnr2L7z06ilSyRhagzErMRARRKA+Gefwm+N85wd/ZmLqIoXH7+Er9/ZSrgb4foSuQV0zqUoplFYUSxW6NjXz9JNf5NbuHM/sHuK5wVdRShFzbSIjGCPYtkXMtRjYf4Rnf/0ym9qa+O6j2/hY5wbKlYCG+jpsW6+BrpmCy2VZmnLFJxF32PnwZ9jcsYE9f3qNkfFpdnz907Q0pRERivNVdv/hEIffHCNmGeaKF9j1s+epVEM2NMTp7dnIA/d9ikTCxchaiPcEWEpJGEaEoSF/zxZubGvkx3sPsPOZfXz7obtI1rn8cM8w0zMlIr9MkMpy0c2R7EiT1pryfInB4Xf451tjbO3tIubaeL63foCllADMzVfp7Wnn6Sfv5Sd7D7Dr5y9iWYrW5gZiVki8ZSO5rh4wBmMMAiSSSVpybZwZOcHv9x1kQ3MTq/zXB7AkS2tK5SqNDUn6H/sce/cdxrJsEk7E7/46x8e7evA9bw10FEWEvk/bTV2Yapm5mXdpb+9cifAHAViC8IKQyBge2n4737q/j9eOjpJt7UBEUKvML5GgtCIIfG7IdRBGhtLM+f8cAEArhTHgBRGTU7NMXSiTymSIougqNVshxuDEE9hunLOTp68NYPllvdKoLi8wtaQuNZVVS+BKAJHFvf1+Q0Tw/ZB0qo5M0qFcKmFpvdiFrpCglCLwqkR+ldZc5/KTNYtQRHAdW+pTjhgRrFr1dpUiIzTUu9zd18XeF0/SnGtT4vuLh4xVEREj2HGX2dGTohGaGpuuBCgWwXEsNTpxQb3wyklV8cKa9X4tMFi2JpFMYVWLnDs9QnPn5jAMfC3GLHJoLU4sbmanJqzZydMqm81S8peb4QrAeLFIwrXl9bfG+ds/Tsl6jyHCYsVMptPyr3eOq8AP7exHWnFcF4DA89XMxIQ+f/o49fVpsSxLed7cagAlADcHztwp27s9VZfUjQ116/FeVrUSKIAoJe7UqaMPXJgY2e7E4hsVSgd+ddKvLOzPpNK7U3U3LJT8GUyYHQWgUDDXcNh6D932C6c5cbbdBJGVjFtnxoYL1atNXQvQ339N2xKAl9EMF8I19/IDFt1vr2yPQsHUBvjwpOjvV5fMhNoHqOsC4N+k3GV9Phk2FgAAAABJRU5ErkJggg=="
        Return d
    End Function

    Public Function GetIcon(cheie As String) As Image
        If Not cache.ContainsKey(cheie) Then
            Dim ms As New MemoryStream(Convert.FromBase64String(date_(cheie)))
            cache(cheie) = Image.FromStream(ms)
        End If
        Return cache(cheie)
    End Function
End Module

' ====================== utilitare / fisiere ======================
Public Module Util
    Public Const SECTIUNE_INI As String = "roMail"          ' sectiunea proprie din setari.ini
    Public ReadOnly BaseDir As String = AppDomain.CurrentDomain.BaseDirectory
    Public ReadOnly PostIni As String = Path.Combine(BaseDir, "post.ini")      ' langa program
    ' folder partajat cu alte aplicatii: %UserProfile%\Documents\roSite\Aplicatii
    Public ReadOnly DirAplicatii As String = Path.Combine(Path.Combine(Path.Combine( _
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents"), "roSite"), "Aplicatii")
    Public ReadOnly SetariIni As String = Path.Combine(DirAplicatii, "setari.ini")
    Public Const FARA_CATEGORIE As String = "fara categorie"

    ' Citeste textul unui fisier si returneaza si encodarea detectata (ca fisierul sa poata fi rescris la fel):
    ' UTF-16, UTF-8 cu BOM, UTF-8 fara BOM, altfel ANSI
    Public Function CitesteTextEnc(cale As String, ByRef enc As Encoding) As String
        Dim b() As Byte = File.ReadAllBytes(cale)
        If b.Length >= 2 AndAlso b(0) = &HFF AndAlso b(1) = &HFE Then
            enc = New UnicodeEncoding(False, True)
            Return Encoding.Unicode.GetString(b, 2, b.Length - 2)
        End If
        If b.Length >= 3 AndAlso b(0) = &HEF AndAlso b(1) = &HBB AndAlso b(2) = &HBF Then
            enc = New UTF8Encoding(True)
            Return New UTF8Encoding(False).GetString(b, 3, b.Length - 3)
        End If
        Try
            Dim t As String = New UTF8Encoding(False, True).GetString(b)
            enc = New UTF8Encoding(False)
            Return t
        Catch
            enc = Encoding.Default
            Return Encoding.Default.GetString(b)
        End Try
    End Function

    Public Function CitesteText(cale As String) As String
        Dim enc As Encoding = Nothing
        Return CitesteTextEnc(cale, enc)
    End Function

    Public Function Linii(text As String) As String()
        Return text.Replace(Convert.ToChar(13).ToString(), "").Split(Convert.ToChar(10))
    End Function

    Public Function CitesteIni(cale As String, sectiune As String) As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If Not File.Exists(cale) Then Return d
        Dim cur As String = ""
        For Each linie As String In Linii(CitesteText(cale))
            Dim l As String = linie.Trim()
            If l.Length = 0 OrElse l.StartsWith(";") OrElse l.StartsWith("#") Then Continue For
            If l.StartsWith("[") AndAlso l.EndsWith("]") Then
                cur = l.Substring(1, l.Length - 2).Trim()
            ElseIf String.Equals(cur, sectiune, StringComparison.OrdinalIgnoreCase) Then
                Dim p As Integer = l.IndexOf("="c)
                If p > 0 Then d(l.Substring(0, p).Trim()) = l.Substring(p + 1).Trim()
            End If
        Next
        Return d
    End Function

    Public Function Valoare(d As Dictionary(Of String, String), cheie As String) As String
        Dim v As String = Nothing
        If d.TryGetValue(cheie, v) Then Return v
        Return ""
    End Function

    ' Pune sau inlocuieste sectiunea [sectiune] din fisierul ini. Restul fisierului (alte sectiuni,
    ' comentarii, encodare, tip de rand nou) ramane neschimbat; daca sectiunea nu exista, se adauga la sfarsit.
    Public Sub IniScrieSectiune(cale As String, sectiune As String, linii As List(Of String))
        Directory.CreateDirectory(Path.GetDirectoryName(cale))
        Dim enc As Encoding = New UTF8Encoding(False)
        Dim text As String = ""
        If File.Exists(cale) Then text = CitesteTextEnc(cale, enc)
        Dim cr As String = Convert.ToChar(13).ToString()
        Dim lf As String = Convert.ToChar(10).ToString()
        Dim nl As String = Environment.NewLine
        If text.Contains(lf) AndAlso Not text.Contains(cr) Then nl = lf
        If text.Contains(cr & lf) Then nl = cr & lf

        Dim existente As New List(Of String)()
        If text.Length > 0 Then existente.AddRange(text.Replace(cr, "").Split(Convert.ToChar(10)))
        Do While existente.Count > 0 AndAlso existente(existente.Count - 1).Trim().Length = 0
            existente.RemoveAt(existente.Count - 1)
        Loop

        Dim start As Integer = -1
        Dim fin As Integer = existente.Count
        For i As Integer = 0 To existente.Count - 1
            Dim t As String = existente(i).Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") Then
                Dim nume As String = t.Substring(1, t.Length - 2).Trim()
                If start < 0 Then
                    If String.Equals(nume, sectiune, StringComparison.OrdinalIgnoreCase) Then start = i
                Else
                    fin = i
                    Exit For
                End If
            End If
        Next

        Dim rez As New List(Of String)()
        If start < 0 Then
            rez.AddRange(existente)
            If rez.Count > 0 Then rez.Add("")
            rez.Add("[" & sectiune & "]")
            rez.AddRange(linii)
        Else
            rez.AddRange(existente.GetRange(0, start))
            rez.Add("[" & sectiune & "]")
            rez.AddRange(linii)
            If fin < existente.Count Then
                rez.Add("")
                rez.AddRange(existente.GetRange(fin, existente.Count - fin))
            End If
        End If
        File.WriteAllText(cale, String.Join(nl, rez.ToArray()) & nl, enc)
    End Sub

    Public Function ContineCI(l As IEnumerable(Of String), a As String) As Boolean
        For Each x As String In l
            If String.Equals(x, a, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    Public Function SplitUnic(s As String, virgula As Boolean) As List(Of String)
        Dim r As New List(Of String)()
        Dim t As String = s
        If virgula Then t = t.Replace(","c, ";"c)
        For Each x As String In t.Split(";"c)
            Dim v As String = x.Trim().Trim(""""c)
            If v.Length > 0 AndAlso Not ContineCI(r, v) Then r.Add(v)
        Next
        Return r
    End Function

    Public Function Lista(lb As ListBox) As List(Of String)
        Dim r As New List(Of String)()
        For Each o As Object In lb.Items
            r.Add(o.ToString())
        Next
        Return r
    End Function
End Module

' ====================== jurnal (romail.log) ======================
Public Module Jurnal
    Public ReadOnly CaleLog As String = Path.Combine(DirAplicatii, "romail.log")
    Private ReadOnly lacat As New Object()

    ' O linie pe actiune: data ora | ACTIUNE | detalii   (parolele nu se scriu niciodata)
    Public Sub Scrie(actiune As String, detalii As String)
        Try
            SyncLock lacat
                Directory.CreateDirectory(DirAplicatii)
                Dim fi As New FileInfo(CaleLog)
                If fi.Exists AndAlso fi.Length > 1048576 Then      ' peste 1 MB: arhiveaza ca romail.log.old
                    Dim veche As String = CaleLog & ".old"
                    If File.Exists(veche) Then File.Delete(veche)
                    File.Move(CaleLog, veche)
                End If
                Dim linie As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | " & actiune.PadRight(18) & " | " & Curata(detalii)
                File.AppendAllText(CaleLog, linie & Environment.NewLine, New UTF8Encoding(False))
            End SyncLock
        Catch
            ' un jurnal care nu se poate scrie nu trebuie sa opreasca programul
        End Try
    End Sub

    Private Function Curata(t As String) As String
        Return t.Replace(Convert.ToChar(13).ToString(), " ").Replace(Convert.ToChar(10).ToString(), " ")
    End Function

    ' Descriere standard a unui email: Subiect (obligatoriu), destinatari, fisiere, expeditor
    Public Function DescrieJob(j As SendJob) As String
        Dim nume As New List(Of String)()
        For Each f As String In j.Fisiere
            nume.Add(Path.GetFileName(f))
        Next
        Dim fis As String = If(nume.Count = 0, "(niciunul)", String.Join("; ", nume.ToArray()))
        Return "Subiect: """ & j.Subiect & """ | Catre: " & String.Join("; ", j.Dest.ToArray()) & _
               " | Fisiere: " & fis & " | Expeditor: " & j.S.Utilizator
    End Function

    Public Function Citeste() As String
        If Not File.Exists(CaleLog) Then Return ""
        Try
            Using fs As New FileStream(CaleLog, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs, Encoding.UTF8, True)
                    Return sr.ReadToEnd()
                End Using
            End Using
        Catch ex As Exception
            Return "Nu pot citi jurnalul: " & ex.Message
        End Try
    End Function

    Public Sub Goleste()
        Try
            File.WriteAllText(CaleLog, "", New UTF8Encoding(False))
        Catch
        End Try
        Scrie("JURNAL GOLIT", "jurnalul a fost sters de utilizator")
    End Sub
End Module

' ====================== setari (expeditor, destinatar, categorii) ======================
' Se pastreaza in [roMail] din setari.ini (Documentele mele\Aplicatii); celelalte sectiuni ale fisierului nu se ating.
Public Class Setari
    Public Host As String = ""
    Public Port As Integer = 587
    Public Securitate As String = "STARTTLS"
    Public Utilizator As String = ""
    Public Parola As String = ""
    Public Destinatar As String = ""
    Public Categorii As List(Of String) = CategoriiImplicite()

    Public Shared Function CategoriiImplicite() As List(Of String)
        Dim l As New List(Of String)()
        l.Add(FARA_CATEGORIE)
        l.Add("anunturi")
        l.Add("consilieri")
        l.Add("adunari")
        l.Add("informare")
        l.Add("important")
        Return l
    End Function

    ' "fara categorie" este mereu prima; fara goluri si duplicate
    Public Shared Function NormalizeazaCategorii(src As IEnumerable(Of String)) As List(Of String)
        Dim r As New List(Of String)()
        r.Add(FARA_CATEGORIE)
        For Each c As String In src
            Dim v As String = c.Trim()
            If v.Length > 0 AndAlso Not ContineCI(r, v) Then r.Add(v)
        Next
        Return r
    End Function

    Public Function EsteComplet() As Boolean
        Return Host.Trim() <> "" AndAlso Utilizator.Trim() <> "" AndAlso Parola <> "" AndAlso Destinatar.Trim() <> ""
    End Function

    Public Shared Function Incarca() As Setari
        Dim s As New Setari()
        Dim d As Dictionary(Of String, String) = CitesteIni(SetariIni, SECTIUNE_INI)
        If Valoare(d, "host") <> "" Then s.Host = Valoare(d, "host")
        Dim p As Integer
        If Integer.TryParse(Valoare(d, "port"), p) Then s.Port = p
        If Valoare(d, "securitate") <> "" Then s.Securitate = Valoare(d, "securitate")
        If Valoare(d, "utilizator") <> "" Then s.Utilizator = Valoare(d, "utilizator")
        If Valoare(d, "parola") <> "" Then
            Try
                s.Parola = Encoding.UTF8.GetString(Convert.FromBase64String(Valoare(d, "parola")))
            Catch
                s.Parola = Valoare(d, "parola")
            End Try
        End If
        If Valoare(d, "destinatar") <> "" Then s.Destinatar = Valoare(d, "destinatar")
        Dim cat As String = Valoare(d, "categorii")
        If cat <> "" Then
            Dim sep As Char = ";"c
            If cat.IndexOf(";"c) < 0 AndAlso cat.IndexOf(","c) >= 0 Then sep = ","c
            s.Categorii = NormalizeazaCategorii(cat.Split(sep))
        End If
        ' SmtpClient nu suporta SSL direct (465); folosim STARTTLS (587)
        If s.Securitate.ToUpperInvariant() = "SSL" Then
            s.Securitate = "STARTTLS"
            If s.Port = 465 Then s.Port = 587
        End If
        Return s
    End Function

    Public Sub Salveaza()
        Dim l As New List(Of String)()
        l.Add("host = " & Host)
        l.Add("port = " & Port.ToString())
        l.Add("securitate = " & Securitate)
        l.Add("utilizator = " & Utilizator)
        l.Add("parola = " & Convert.ToBase64String(Encoding.UTF8.GetBytes(Parola)))
        l.Add("destinatar = " & Destinatar)
        l.Add("categorii = " & String.Join(";", Categorii.ToArray()))
        IniScrieSectiune(SetariIni, SECTIUNE_INI, l)
    End Sub
End Class

' ====================== trimitere ======================
Public Class SendJob
    Public S As Setari
    Public Dest As List(Of String)
    Public Subiect As String
    Public Mesaj As String
    Public Fisiere As List(Of String)
End Class

Public Module Mailer
    Public Sub Trimite(j As SendJob)
        Try
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or CType(3072, SecurityProtocolType)
        Catch
        End Try
        Dim msg As New MailMessage()
        Try
            msg.From = New MailAddress(j.S.Utilizator)
            For Each a As String In j.Dest
                msg.To.Add(a)
            Next
            msg.Subject = j.Subiect
            msg.SubjectEncoding = Encoding.UTF8
            msg.Body = j.Mesaj
            msg.BodyEncoding = Encoding.UTF8
            For Each f As String In j.Fisiere
                msg.Attachments.Add(New Attachment(f))
            Next
            Using smtp As New SmtpClient(j.S.Host, j.S.Port)
                smtp.EnableSsl = (j.S.Securitate.ToUpperInvariant() = "STARTTLS")
                smtp.Credentials = New NetworkCredential(j.S.Utilizator, j.S.Parola)
                smtp.Timeout = 60000
                smtp.Send(msg)
            End Using
        Finally
            msg.Dispose()
        End Try
    End Sub
End Module

' ====================== fereastra principala ======================
Public Class MainForm
    Inherits Form

    Private Shared ReadOnly BG As Color = Color.FromArgb(155, 179, 211)
    Private Const W As Integer = 545        ' latimea zonei client la pornire
    Private Const FW As Integer = W - 28    ' latime control pe tot randul
    Private Const LW As Integer = W - 62    ' latime liste (langa coloana de butoane)
    Private Const RX As Integer = W - 40    ' pozitia X a coloanei de butoane
    Private Const TAG_PUBLIC As String = "public"

    Private txtSub As TextBox
    Private chkPublic As CheckBox
    Private lstCat As ComboBox
    Private txtMsg As TextBox
    Private lstFiles As ListBox
    Private btnSend As Button
    Private btnSet As Button
    Private btnExit As Button
    Private lblStatus As Label
    Private lblReset As Label
    Private lblJurnal As Label
    Private ReadOnly tt As New ToolTip()
    Private cfg As Setari
    Private jobCurent As SendJob
    Private sincron As Boolean = False      ' True = ignora evenimentele (actualizare din cod)

    Public Sub New()
        Me.Text = "roMail"
        Me.BackColor = BG
        Me.Font = New Font("Microsoft Sans Serif", 9.0F)
        Me.ClientSize = New Size(W, 612)
        Me.MinimumSize = New Size(480, 600)
        Me.StartPosition = FormStartPosition.Manual   ' pozitia se seteaza in LaIncarcare (dreapta-sus)
        Try
            Me.Icon = Icon.FromHandle(DirectCast(Icoane.GetIcon("app"), Bitmap).GetHicon())
        Catch
        End Try

        cfg = Setari.Incarca()

        Construieste()
        IncarcaCategorii()
        Jurnal.Scrie("PORNIRE", "roMail pornit | server: " & cfg.Host & ":" & cfg.Port.ToString() & " | expeditor: " & cfg.Utilizator & " | destinatar: " & cfg.Destinatar)
        Reinitializare()
        LayoutJos()
        AddHandler Me.Resize, AddressOf LaRedimensionare
        AddHandler Me.Shown, AddressOf LaAfisare
        AddHandler Me.Load, AddressOf LaIncarcare
        AddHandler Me.FormClosed, AddressOf LaInchis
    End Sub

    ' ---------- construire interfata ----------
    Private Function Eticheta(text As String, x As Integer, y As Integer) As Label
        Dim l As New Label()
        l.Text = text
        l.AutoSize = True
        l.BackColor = Color.Transparent
        l.Location = New Point(x, y)
        Me.Controls.Add(l)
        Return l
    End Function

    Private Function ButonIcon(cheie As String, x As Integer, y As Integer, sfat As String, chenar As Boolean, ancora As AnchorStyles) As Button
        Dim b As New Button()
        b.Image = Icoane.GetIcon(cheie)
        b.SetBounds(x, y, 26, 26)
        b.Anchor = ancora
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = If(chenar, 1, 0)
        b.FlatAppearance.BorderColor = Color.FromArgb(60, 60, 60)
        b.BackColor = If(chenar, Color.FromArgb(233, 239, 247), BG)
        b.UseVisualStyleBackColor = False
        b.Cursor = Cursors.Hand
        tt.SetToolTip(b, sfat)
        Me.Controls.Add(b)
        Return b
    End Function

    Private Sub Construieste()
        Dim fataSt As AnchorStyles = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Dim jos As AnchorStyles = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Dim dreaptaJos As AnchorStyles = AnchorStyles.Bottom Or AnchorStyles.Right
        Dim fontMare As New Font("Microsoft Sans Serif", 10.0F)
        Dim fontB As New Font("Microsoft Sans Serif", 9.0F, FontStyle.Bold)

        ' Subiect
        Eticheta("Subiect:", 14, 8)
        Dim pnl As New Panel()
        pnl.BackColor = Color.White
        pnl.SetBounds(14, 26, FW, 32)
        pnl.Anchor = fataSt
        txtSub = New TextBox()
        txtSub.BorderStyle = BorderStyle.None
        txtSub.BackColor = Color.White
        txtSub.Font = fontMare
        txtSub.SetBounds(5, 7, FW - 10, 20)
        txtSub.Anchor = fataSt
        Dim ul As New Panel()
        ul.BackColor = Color.FromArgb(47, 109, 181)
        ul.SetBounds(0, 30, FW, 2)
        ul.Anchor = fataSt
        pnl.Controls.Add(txtSub)
        pnl.Controls.Add(ul)
        Me.Controls.Add(pnl)
        AddHandler txtSub.TextChanged, AddressOf LaSubiectModificat

        ' Publicare pe prima pagina
        chkPublic = New CheckBox()
        chkPublic.Text = "Publicare pe prima pagina"
        chkPublic.AutoSize = True
        chkPublic.BackColor = Color.Transparent
        chkPublic.Font = fontB
        chkPublic.Location = New Point(14, 68)
        chkPublic.Cursor = Cursors.Hand
        AddHandler chkPublic.CheckedChanged, AddressOf LaPublicare
        Me.Controls.Add(chkPublic)

        ' Categorie (selectie unica, din setari.ini) - pe acelasi rand cu bifa
        Eticheta("Categorie:", 222, 70)
        lstCat = New ComboBox()
        lstCat.DropDownStyle = ComboBoxStyle.DropDownList
        lstCat.SetBounds(290, 65, W - 14 - 290, 24)
        lstCat.Anchor = fataSt
        lstCat.Font = fontMare
        lstCat.BackColor = Color.White
        lstCat.MaxDropDownItems = 12
        lstCat.Cursor = Cursors.Hand
        AddHandler lstCat.SelectedIndexChanged, AddressOf LaCategorie
        Me.Controls.Add(lstCat)

        ' Mesaj
        Eticheta("Mesaj:", 14, 100)
        txtMsg = New TextBox()
        txtMsg.SetBounds(14, 118, FW, 274)
        txtMsg.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        txtMsg.Font = fontMare
        txtMsg.Multiline = True
        txtMsg.ScrollBars = ScrollBars.Vertical
        txtMsg.AcceptsReturn = True
        txtMsg.MaxLength = Int32.MaxValue
        txtMsg.BorderStyle = BorderStyle.Fixed3D
        txtMsg.BackColor = Color.White
        Me.Controls.Add(txtMsg)

        ' Fisiere atasate
        Eticheta("Fisiere atasate:", 14, 400)
        Dim ind As Label = Eticheta("(dublu clic pentru vizualizare)", 112, 402)
        ind.Font = New Font("Microsoft Sans Serif", 8.0F)
        lstFiles = New ListBox()
        lstFiles.SetBounds(14, 418, LW, 84)
        lstFiles.Anchor = jos
        lstFiles.SelectionMode = SelectionMode.MultiExtended
        lstFiles.IntegralHeight = False
        lstFiles.BackColor = Color.White
        lstFiles.HorizontalScrollBar = True
        AddHandler lstFiles.DoubleClick, AddressOf DeschideFisier
        AddHandler lstFiles.KeyDown, AddressOf TastaFisiere
        Me.Controls.Add(lstFiles)
        Dim b As Button
        b = ButonIcon("plus", RX, 418, "Adauga fisiere", False, dreaptaJos)
        AddHandler b.Click, AddressOf AdaugaFisiere
        b = ButonIcon("trash", RX, 446, "Sterge toate fisierele din lista", True, dreaptaJos)
        AddHandler b.Click, AddressOf GolesteFisiere
        b = ButonIcon("redx", RX, 474, "Scoate fisierul selectat", False, dreaptaJos)
        AddHandler b.Click, AddressOf ScoateFisier

        ' Butoane jos (pozitionate in LayoutJos)
        Dim fontBtn As New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnSend = New Button()
        btnSend.Text = "Trimite &Mail"
        btnSend.Image = Icoane.GetIcon("env")
        btnSend.ImageAlign = ContentAlignment.MiddleLeft
        btnSend.TextAlign = ContentAlignment.MiddleCenter
        btnSend.Font = fontBtn
        btnSend.BackColor = Color.FromArgb(240, 240, 240)
        btnSend.UseVisualStyleBackColor = False
        btnSend.Cursor = Cursors.Hand
        AddHandler btnSend.Click, AddressOf Trimite
        Me.Controls.Add(btnSend)

        btnSet = New Button()
        btnSet.Image = Icoane.GetIcon("floppy")
        btnSet.FlatStyle = FlatStyle.Flat
        btnSet.FlatAppearance.BorderSize = 0
        btnSet.BackColor = Color.FromArgb(166, 166, 166)
        btnSet.UseVisualStyleBackColor = False
        btnSet.Cursor = Cursors.Hand
        tt.SetToolTip(btnSet, "Setari (expeditor, destinatar, categorii)")
        AddHandler btnSet.Click, AddressOf DeschideSetari
        Me.Controls.Add(btnSet)

        btnExit = New Button()
        btnExit.Text = "&Iesire"
        btnExit.Image = Icoane.GetIcon("stop")
        btnExit.ImageAlign = ContentAlignment.MiddleLeft
        btnExit.TextAlign = ContentAlignment.MiddleCenter
        btnExit.Font = fontBtn
        btnExit.BackColor = Color.FromArgb(255, 215, 0)
        btnExit.UseVisualStyleBackColor = False
        btnExit.Cursor = Cursors.Hand
        AddHandler btnExit.Click, AddressOf Iesire
        Me.Controls.Add(btnExit)

        ' Bara de stare
        Dim bara As New Panel()
        bara.Dock = DockStyle.Bottom
        bara.Height = 22
        bara.BackColor = Color.FromArgb(237, 237, 237)
        lblReset = New Label()
        lblReset.Text = "Reinitializare"
        lblReset.ForeColor = Color.Red
        lblReset.Font = fontB
        lblReset.AutoSize = True
        lblReset.Cursor = Cursors.Hand
        lblReset.Location = New Point(14, 4)
        AddHandler lblReset.Click, AddressOf LaReinitializare
        lblStatus = New Label()
        lblStatus.AutoSize = False
        lblStatus.AutoEllipsis = True
        lblStatus.SetBounds(120, 4, 340, 16)
        lblJurnal = New Label()
        lblJurnal.Text = "Jurnal"
        lblJurnal.ForeColor = Color.FromArgb(0, 0, 200)
        lblJurnal.Font = New Font("Microsoft Sans Serif", 9.0F, FontStyle.Underline)
        lblJurnal.AutoSize = True
        lblJurnal.Cursor = Cursors.Hand
        lblJurnal.Location = New Point(W - 62, 4)   ' repozitionat in LayoutJos
        tt.SetToolTip(lblJurnal, "Deschide jurnalul actiunilor (romail.log)")
        AddHandler lblJurnal.Click, AddressOf DeschideJurnal
        bara.Controls.Add(lblReset)
        bara.Controls.Add(lblStatus)
        bara.Controls.Add(lblJurnal)
        Me.Controls.Add(bara)
    End Sub

    Private Sub LayoutJos()
        If btnSend Is Nothing Then Return
        Dim w As Integer = Me.ClientSize.Width
        Dim y As Integer = Me.ClientSize.Height - 22 - 66
        Dim lat As Integer = (w - 40) \ 2 - 28
        btnSend.SetBounds(14, y, lat, 48)
        btnSet.SetBounds((w - 40) \ 2, y + 4, 40, 40)
        btnExit.SetBounds((w + 40) \ 2 + 14, y, lat, 48)
        If lblJurnal IsNot Nothing AndAlso lblJurnal.Parent IsNot Nothing Then
            lblJurnal.Left = lblJurnal.Parent.ClientSize.Width - lblJurnal.Width - 12
        End If
    End Sub

    Private Sub LaRedimensionare(sender As Object, e As EventArgs)
        LayoutJos()
    End Sub

    Private Sub LaIncarcare(sender As Object, e As EventArgs)
        ' coltul din dreapta-sus al zonei de lucru a ecranului principal
        Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Location = New Point(wa.Right - Me.Width, wa.Top)
    End Sub

    Private Sub LaInchis(sender As Object, e As FormClosedEventArgs)
        Jurnal.Scrie("INCHIDERE", "roMail inchis")
    End Sub

    Private Sub LaAfisare(sender As Object, e As EventArgs)
        LayoutJos()
        txtSub.Focus()
        txtSub.SelectionStart = txtSub.Text.Length
        If Not cfg.EsteComplet() Then
            ' prima rulare (sau setari lipsa): se cer datele, nu exista valori implicite in cod
            Jurnal.Scrie("SETARI NECOMPLETATE", "lipsesc setarile din [" & SECTIUNE_INI & "] in " & SetariIni)
            MessageBox.Show(Me, "Setarile nu sunt completate (server, utilizator, parola, destinatar)." & Environment.NewLine & _
                            "Se deschide fereastra Setari.", "roMail", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DeschideSetari(Me, EventArgs.Empty)
        End If
    End Sub

    ' ---------- stare ----------
    Private Sub Stare(text As String)
        lblStatus.Text = text
    End Sub

    Private Sub LaReinitializare(sender As Object, e As EventArgs)
        Reinitializare(True)
    End Sub

    Private Sub Reinitializare(Optional manual As Boolean = False)
        lstFiles.Items.Clear()
        Dim p As Dictionary(Of String, String) = CitesteIni(PostIni, "Email")
        txtSub.Text = Valoare(p, "Sub")
        txtMsg.Text = Valoare(p, "Msg").Replace("\n", Environment.NewLine)
        For Each f As String In SplitUnic(Valoare(p, "File"), False)
            lstFiles.Items.Add(f)
        Next
        SincronizeazaDinSubiect()
        If Not cfg.EsteComplet() Then
            Stare("Setari necompletate - butonul cu discheta")
        ElseIf p.Count > 0 Then
            Stare("post.ini incarcat | Catre: " & cfg.Destinatar)
        Else
            Stare("Catre: " & cfg.Destinatar)
        End If
        Jurnal.Scrie(If(manual, "REINITIALIZARE", "DATE INITIALE"), "post.ini: " & If(File.Exists(PostIni), "gasit", "negasit") & _
                     " | Subiect: """ & txtSub.Text & """ | fisiere: " & lstFiles.Items.Count.ToString())
        txtSub.Focus()
        txtSub.SelectionStart = txtSub.Text.Length
    End Sub

    ' ---------- Publicare / categorie -> etichete in fata subiectului ----------
    Private Function EsteCategorie(nume As String) As Boolean
        If String.Equals(nume, FARA_CATEGORIE, StringComparison.OrdinalIgnoreCase) Then Return False
        Return ContineCI(cfg.Categorii, nume)
    End Function

    Private Function NumeCanonic(nume As String) As String
        For Each c As String In cfg.Categorii
            If String.Equals(c, nume, StringComparison.OrdinalIgnoreCase) Then Return c
        Next
        Return nume
    End Function

    ' Scoate din fata subiectului etichetele cunoscute: [public] si [categorie din lista]
    Private Sub DespartTaguri(subiect As String, ByRef publicare As Boolean, ByRef categorie As String, ByRef rest As String)
        publicare = False
        categorie = ""
        Dim t As String = subiect.TrimStart()
        Do While t.StartsWith("[")
            Dim j As Integer = t.IndexOf("]"c)
            If j < 0 Then Exit Do
            Dim nume As String = t.Substring(1, j - 1).Trim()
            If String.Equals(nume, TAG_PUBLIC, StringComparison.OrdinalIgnoreCase) Then
                publicare = True
            ElseIf EsteCategorie(nume) Then
                If categorie = "" Then categorie = NumeCanonic(nume)
            Else
                Exit Do
            End If
            t = t.Substring(j + 1).TrimStart()
        Loop
        rest = t
    End Sub

    Private Function CategorieSelectata() As String
        If lstCat.SelectedIndex < 0 Then Return ""
        Dim n As String = lstCat.SelectedItem.ToString()
        If String.Equals(n, FARA_CATEGORIE, StringComparison.OrdinalIgnoreCase) Then Return ""
        Return n
    End Function

    Private Sub SelecteazaCategorie(nume As String)
        For i As Integer = 0 To lstCat.Items.Count - 1
            If String.Equals(lstCat.Items(i).ToString(), nume, StringComparison.OrdinalIgnoreCase) Then
                lstCat.SelectedIndex = i
                Return
            End If
        Next
        If lstCat.Items.Count > 0 Then lstCat.SelectedIndex = 0
    End Sub

    Private Sub IncarcaCategorii()
        Dim anterior As String = CategorieSelectata()
        sincron = True
        Try
            lstCat.Items.Clear()
            For Each c As String In cfg.Categorii
                lstCat.Items.Add(c)
            Next
            SelecteazaCategorie(If(anterior = "", FARA_CATEGORIE, anterior))
        Finally
            sincron = False
        End Try
        SincronizeazaDinSubiect()
    End Sub

    ' bifa / categorie schimbate de utilizator -> rescrie etichetele din fata subiectului
    Private Sub AplicaTaguri()
        If sincron Then Return
        Dim pub As Boolean
        Dim cat As String = ""
        Dim rest As String = ""
        DespartTaguri(txtSub.Text, pub, cat, rest)
        Dim prefix As String = ""
        If chkPublic.Checked Then prefix &= "[" & TAG_PUBLIC & "]"
        Dim sel As String = CategorieSelectata()
        If sel <> "" Then prefix &= "[" & sel & "]"
        Dim nou As String = prefix & rest
        If prefix <> "" AndAlso rest <> "" Then nou = prefix & " " & rest
        sincron = True
        Try
            txtSub.Text = nou
            txtSub.SelectionStart = txtSub.Text.Length
        Finally
            sincron = False
        End Try
    End Sub

    ' subiect modificat (tastat sau din post.ini) -> bifa si categoria reflecta etichetele din el
    Private Sub SincronizeazaDinSubiect()
        If sincron Then Return
        sincron = True
        Try
            Dim pub As Boolean
            Dim cat As String = ""
            Dim rest As String = ""
            DespartTaguri(txtSub.Text, pub, cat, rest)
            chkPublic.Checked = pub
            SelecteazaCategorie(If(cat = "", FARA_CATEGORIE, cat))
        Finally
            sincron = False
        End Try
    End Sub

    Private Sub LaPublicare(sender As Object, e As EventArgs)
        AplicaTaguri()
    End Sub

    Private Sub LaCategorie(sender As Object, e As EventArgs)
        AplicaTaguri()
    End Sub

    Private Sub LaSubiectModificat(sender As Object, e As EventArgs)
        SincronizeazaDinSubiect()
    End Sub

    ' ---------- fisiere ----------
    Private Sub AdaugaFisiere(sender As Object, e As EventArgs)
        Using d As New OpenFileDialog()
            d.Title = "Alege fisiere de atasat"
            d.Multiselect = True
            If d.ShowDialog(Me) = DialogResult.OK Then
                For Each f As String In d.FileNames
                    If Not ContineCI(Lista(lstFiles), f) Then
                        lstFiles.Items.Add(f)
                        Jurnal.Scrie("FISIER ADAUGAT", f)
                    End If
                Next
            End If
        End Using
    End Sub

    Private Sub ScoateFisier(sender As Object, e As EventArgs)
        For i As Integer = lstFiles.Items.Count - 1 To 0 Step -1
            If lstFiles.GetSelected(i) Then
                Jurnal.Scrie("FISIER SCOS", lstFiles.Items(i).ToString())
                lstFiles.Items.RemoveAt(i)
            End If
        Next
    End Sub

    Private Sub TastaFisiere(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Delete Then ScoateFisier(sender, EventArgs.Empty)
    End Sub

    Private Sub GolesteFisiere(sender As Object, e As EventArgs)
        If lstFiles.Items.Count > 0 Then
            If MessageBox.Show(Me, "Scoateti toate fisierele din lista?", "Fisiere atasate", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Jurnal.Scrie("FISIERE GOLITE", lstFiles.Items.Count.ToString() & " fisiere scoase din lista")
                lstFiles.Items.Clear()
            End If
        End If
    End Sub

    Private Sub DeschideFisier(sender As Object, e As EventArgs)
        If lstFiles.SelectedIndex < 0 Then Return
        Try
            Dim psi As New Diagnostics.ProcessStartInfo(lstFiles.SelectedItem.ToString())
            psi.UseShellExecute = True
            Diagnostics.Process.Start(psi)
            Jurnal.Scrie("FISIER DESCHIS", lstFiles.SelectedItem.ToString())
        Catch ex As Exception
            MessageBox.Show(Me, "Nu pot deschide fisierul:" & Environment.NewLine & ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ---------- setari ----------
    Private Sub DeschideSetari(sender As Object, e As EventArgs)
        Using f As New SettingsForm(cfg)
            If f.ShowDialog(Me) = DialogResult.OK Then
                cfg = f.Rezultat          ' testat si deja salvat in setari.ini
                IncarcaCategorii()
                Stare("Setari salvate | Catre: " & cfg.Destinatar)
            End If
        End Using
    End Sub

    Private Sub DeschideJurnal(sender As Object, e As EventArgs)
        Using f As New LogForm(cfg)
            f.ShowDialog(Me)
        End Using
    End Sub

    Private Sub Iesire(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ---------- trimitere ----------
    Private Sub Trimite(sender As Object, e As EventArgs)
        If Not btnSend.Enabled Then Return
        Dim dest As String = cfg.Destinatar.Trim()
        If Not cfg.EsteComplet() OrElse Not dest.Contains("@") Then
            Jurnal.Scrie("TRIMITERE BLOCATA", "Subiect: """ & txtSub.Text & """ | motiv: setari necompletate")
            MessageBox.Show(Me, "Setarile nu sunt completate (server, utilizator, parola, destinatar)." & Environment.NewLine & _
                            "Completati-le in Setari (butonul cu discheta).", "Setari necompletate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim fisiere As List(Of String) = Lista(lstFiles)
        Dim lipsa As String = ""
        For Each f As String In fisiere
            If Not File.Exists(f) Then lipsa &= f & Environment.NewLine
        Next
        If lipsa <> "" Then
            Jurnal.Scrie("TRIMITERE BLOCATA", "Subiect: """ & txtSub.Text & """ | motiv: fisier negasit: " & lipsa.Replace(Environment.NewLine, "; "))
            MessageBox.Show(Me, "Nu exista:" & Environment.NewLine & lipsa, "Fisier negasit", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim job As New SendJob()
        job.S = cfg
        job.Dest = New List(Of String)()
        job.Dest.Add(dest)
        job.Subiect = txtSub.Text
        job.Mesaj = txtMsg.Text
        job.Fisiere = fisiere
        jobCurent = job

        btnSend.Enabled = False
        Stare("Se trimite...")
        Me.Cursor = Cursors.WaitCursor
        Dim bw As New BackgroundWorker()
        AddHandler bw.DoWork, AddressOf LucruTrimitere
        AddHandler bw.RunWorkerCompleted, AddressOf FinalTrimitere
        bw.RunWorkerAsync(job)
    End Sub

    Private Sub LucruTrimitere(sender As Object, e As DoWorkEventArgs)
        Mailer.Trimite(DirectCast(e.Argument, SendJob))
        e.Result = "ok"
    End Sub

    Private Sub FinalTrimitere(sender As Object, e As RunWorkerCompletedEventArgs)
        Me.Cursor = Cursors.Default
        btnSend.Enabled = True
        If e.Error Is Nothing Then
            Stare("Email trimis.")
            Jurnal.Scrie("TRIMITERE OK", Jurnal.DescrieJob(jobCurent))
            MessageBox.Show(Me, "Email trimis cu succes:" & Environment.NewLine & jobCurent.Subiect & Environment.NewLine & "Catre: " & String.Join("; ", jobCurent.Dest.ToArray()), "Rezultat trimitere", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Stare("Eroare la trimitere.")
            Dim t As String = e.Error.GetType().Name & ": " & e.Error.Message
            If e.Error.InnerException IsNot Nothing Then t &= Environment.NewLine & e.Error.InnerException.Message
            Jurnal.Scrie("TRIMITERE ESUATA", Jurnal.DescrieJob(jobCurent) & " | Eroare: " & t)
            MessageBox.Show(Me, "Trimiterea a esuat:" & Environment.NewLine & t, "Rezultat trimitere", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class

' ====================== dialog: setari ======================
' La "Salveaza": se trimite un email de test catre adresa expeditorului;
' setari.ini se scrie DOAR daca testul reuseste.
Public Class SettingsForm
    Inherits Form
    Private ReadOnly txtHost As New TextBox()
    Private ReadOnly txtPort As New TextBox()
    Private ReadOnly cmbSec As New ComboBox()
    Private ReadOnly txtUser As New TextBox()
    Private ReadOnly txtPass As New TextBox()
    Private ReadOnly txtDest As New TextBox()
    Private ReadOnly txtCat As New TextBox()
    Private ReadOnly lblTest As New Label()
    Private ReadOnly btnOk As New Button()
    Private ReadOnly btnRenunta As New Button()
    Private testInCurs As Boolean = False
    Private Const TEST_SUBIECT As String = "roMail - test setari expeditor"
    Private candidat As Setari
    Public Rezultat As Setari

    Public Sub New(s As Setari)
        Me.Text = "Setari"
        Me.BackColor = Color.FromArgb(155, 179, 211)
        Me.Font = New Font("Microsoft Sans Serif", 9.0F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(420, 392)

        Rand("Server SMTP:", txtHost, 14)
        Rand("Port:", txtPort, 46)
        Rand("Securitate:", cmbSec, 78)
        Rand("Utilizator (email):", txtUser, 110)
        Rand("Parola:", txtPass, 142)
        Rand("Destinatar (unic):", txtDest, 182)
        cmbSec.DropDownStyle = ComboBoxStyle.DropDownList
        cmbSec.Items.Add("STARTTLS")
        cmbSec.Items.Add("Fara criptare")
        txtPass.UseSystemPasswordChar = True

        Dim lc As New Label()
        lc.Text = "Categorii:" & Environment.NewLine & "(una pe linie)"
        lc.AutoSize = True
        lc.Location = New Point(12, 218)
        Me.Controls.Add(lc)
        txtCat.SetBounds(140, 214, 268, 100)
        txtCat.Multiline = True
        txtCat.ScrollBars = ScrollBars.Vertical
        txtCat.AcceptsReturn = True
        Me.Controls.Add(txtCat)

        txtHost.Text = s.Host
        txtPort.Text = s.Port.ToString()
        cmbSec.SelectedIndex = If(s.Securitate.ToUpperInvariant() = "STARTTLS", 0, 1)
        txtUser.Text = s.Utilizator
        txtPass.Text = s.Parola
        txtDest.Text = s.Destinatar
        txtCat.Text = String.Join(Environment.NewLine, s.Categorii.ToArray())

        lblTest.AutoSize = False
        lblTest.SetBounds(12, 322, 396, 18)
        Me.Controls.Add(lblTest)

        btnOk.Text = "Salveaza"
        btnOk.SetBounds(100, 352, 100, 28)
        AddHandler btnOk.Click, AddressOf LaSalveaza
        btnRenunta.Text = "Renunta"
        btnRenunta.SetBounds(220, 352, 100, 28)
        btnRenunta.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnOk)
        Me.Controls.Add(btnRenunta)
        Me.AcceptButton = btnOk
        Me.CancelButton = btnRenunta
    End Sub

    Private Sub Rand(text As String, c As Control, y As Integer)
        Dim l As New Label()
        l.Text = text
        l.AutoSize = True
        l.Location = New Point(12, y + 4)
        Me.Controls.Add(l)
        c.SetBounds(140, y, 268, 24)
        Me.Controls.Add(c)
    End Sub

    Private Sub BlocheazaUI(blocat As Boolean)
        testInCurs = blocat
        For Each c As Control In Me.Controls
            If Not Object.ReferenceEquals(c, lblTest) Then c.Enabled = Not blocat
        Next
        Me.Cursor = If(blocat, Cursors.WaitCursor, Cursors.Default)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If testInCurs Then e.Cancel = True   ' nu se inchide cat timp ruleaza testul
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub LaSalveaza(sender As Object, e As EventArgs)
        Dim p As Integer
        If Not Integer.TryParse(txtPort.Text.Trim(), p) OrElse p <= 0 OrElse p > 65535 Then
            MessageBox.Show(Me, "Portul trebuie sa fie un numar valid.", "Port invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If txtHost.Text.Trim() = "" Then
            MessageBox.Show(Me, "Completati serverul SMTP.", "Server lipsa", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim adr As New MailAddress(txtUser.Text.Trim())
        Catch
            MessageBox.Show(Me, "Adresa de email a expeditorului nu este valida.", "Utilizator invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try
        Try
            Dim adr2 As New MailAddress(txtDest.Text.Trim())
        Catch
            MessageBox.Show(Me, "Adresa destinatarului nu este valida.", "Destinatar invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try
        Dim brute As New List(Of String)()
        For Each linie As String In Linii(txtCat.Text)
            Dim v As String = linie.Trim()
            If v.Length = 0 Then Continue For
            If v.IndexOfAny(New Char() {"["c, "]"c, ";"c}) >= 0 Then
                MessageBox.Show(Me, "Categoria """ & v & """ contine caractere nepermise ( [ ] ; ).", "Categorie invalida", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            brute.Add(v)
        Next

        Dim r As New Setari()
        r.Host = txtHost.Text.Trim()
        r.Port = p
        r.Securitate = If(cmbSec.SelectedIndex = 0, "STARTTLS", "NONE")
        r.Utilizator = txtUser.Text.Trim()
        r.Parola = txtPass.Text
        r.Destinatar = txtDest.Text.Trim()
        r.Categorii = Setari.NormalizeazaCategorii(brute)
        candidat = r
        Jurnal.Scrie("SETARI TEST", "server: " & r.Host & ":" & r.Port.ToString() & " (" & r.Securitate & ") | utilizator: " & r.Utilizator & _
                     " | Subiect: """ & TEST_SUBIECT & """")

        BlocheazaUI(True)
        lblTest.ForeColor = Color.FromArgb(30, 30, 30)
        lblTest.Text = "Se verifica setarile (email de test catre expeditor)..."
        Dim bw As New BackgroundWorker()
        AddHandler bw.DoWork, AddressOf LucruTest
        AddHandler bw.RunWorkerCompleted, AddressOf FinalTest
        bw.RunWorkerAsync(r)
    End Sub

    Private Sub LucruTest(sender As Object, e As DoWorkEventArgs)
        Dim s As Setari = DirectCast(e.Argument, Setari)
        Dim j As New SendJob()
        j.S = s
        j.Dest = New List(Of String)()
        j.Dest.Add(s.Utilizator)
        j.Subiect = TEST_SUBIECT
        j.Mesaj = "Mesaj de test trimis de roMail pentru verificarea setarilor." & Environment.NewLine & _
                  "Daca l-ati primit, setarile sunt corecte."
        j.Fisiere = New List(Of String)()
        Mailer.Trimite(j)
    End Sub

    Private Sub FinalTest(sender As Object, e As RunWorkerCompletedEventArgs)
        BlocheazaUI(False)
        If e.Error IsNot Nothing Then
            lblTest.ForeColor = Color.DarkRed
            lblTest.Text = "Test esuat - setarile NU au fost salvate."
            Dim t As String = e.Error.GetType().Name & ": " & e.Error.Message
            If e.Error.InnerException IsNot Nothing Then t &= Environment.NewLine & e.Error.InnerException.Message
            Jurnal.Scrie("SETARI ESUATE", "test esuat, setari.ini neschimbat | Subiect: """ & TEST_SUBIECT & """ | utilizator: " & candidat.Utilizator & " | Eroare: " & t)
            MessageBox.Show(Me, "Testul de trimitere a esuat, setarile nu au fost salvate." & Environment.NewLine & Environment.NewLine & t, _
                            "Verificare setari", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Try
            candidat.Salveaza()
        Catch ex As Exception
            lblTest.ForeColor = Color.DarkRed
            lblTest.Text = "Testul a reusit, dar setari.ini nu a putut fi scris."
            Jurnal.Scrie("SETARI ESUATE", "test reusit, dar setari.ini nu a putut fi scris: " & ex.Message)
            MessageBox.Show(Me, "Nu pot scrie setari.ini:" & Environment.NewLine & ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        Rezultat = candidat
        Jurnal.Scrie("SETARI SALVATE", "test reusit, setari.ini actualizat | Subiect: """ & TEST_SUBIECT & """ | server: " & candidat.Host & ":" & candidat.Port.ToString() & _
                     " | utilizator: " & candidat.Utilizator & " | destinatar: " & candidat.Destinatar & " | categorii: " & String.Join("; ", candidat.Categorii.ToArray()))
        MessageBox.Show(Me, "Testul a reusit: email trimis catre " & candidat.Utilizator & "." & Environment.NewLine & "Setarile au fost salvate.", _
                        "Verificare setari", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.DialogResult = DialogResult.OK
    End Sub
End Class

' ====================== fereastra: jurnal ======================
' Lista liniilor din romail.log. O linie "TRIMITERE OK" poate fi selectata, iar butonul
' "Anulare publicare" trimite imediat, catre aceiasi destinatari, un email cu Subiectul
' "[sterg]" + subiectul original.
Public Class LogForm
    Inherits Form
    Private Const PREFIX_STERG As String = "[sterg]"
    Private Const TEXT_ANULARE As String = "Anulare publicare."
    Private Const CONFIRMA_ANULARE As Boolean = True    ' False = trimite fara intrebare

    Private ReadOnly lst As New ListBox()
    Private ReadOnly lblSel As New Label()
    Private ReadOnly btnAnulare As New Button()
    Private ReadOnly cfg As Setari
    Private selSubiect As String = ""
    Private selCatre As New List(Of String)()
    Private jobAnulare As SendJob
    Private trimitere As Boolean = False

    Public Sub New(s As Setari)
        cfg = s
        Me.Text = "Jurnal roMail (romail.log)"
        Me.BackColor = Color.FromArgb(155, 179, 211)
        Me.Font = New Font("Microsoft Sans Serif", 9.0F)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ShowInTaskbar = False
        Me.ClientSize = New Size(760, 470)
        Me.MinimumSize = New Size(560, 320)

        lst.SetBounds(10, 10, 740, 376)
        lst.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        lst.Font = New Font("Consolas", 9.0F)
        lst.IntegralHeight = False
        lst.HorizontalScrollBar = True
        lst.SelectionMode = SelectionMode.One
        lst.BackColor = Color.White
        AddHandler lst.SelectedIndexChanged, AddressOf LaSelectie
        Me.Controls.Add(lst)

        lblSel.SetBounds(10, 392, 740, 20)
        lblSel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        lblSel.AutoEllipsis = True
        Me.Controls.Add(lblSel)

        Dim jos As AnchorStyles = AnchorStyles.Bottom Or AnchorStyles.Left
        Buton("Reimprospateaza", 10, 120, jos, AddressOf LaReimprospateaza)
        Buton("Deschide in Notepad", 138, 140, jos, AddressOf LaNotepad)
        Buton("Goleste jurnalul", 286, 120, jos, AddressOf LaGoleste)
        btnAnulare = Buton("Anulare publicare", 414, 140, jos, AddressOf LaAnulare)
        Buton("Inchide", 630, 120, AnchorStyles.Bottom Or AnchorStyles.Right, AddressOf LaInchide)
        AddHandler Me.Shown, AddressOf LaAfisare
    End Sub

    Private Function Buton(text As String, x As Integer, w As Integer, ancora As AnchorStyles, h As EventHandler) As Button
        Dim b As New Button()
        b.Text = text
        b.SetBounds(x, 426, w, 28)
        b.Anchor = ancora
        AddHandler b.Click, h
        Me.Controls.Add(b)
        Return b
    End Function

    ' --- intrare "TRIMITERE OK": extrage Subiectul si destinatarii din linia de jurnal ---
    Private Shared Function ParseOk(linie As String, ByRef subiect As String, ByRef catre As List(Of String)) As Boolean
        subiect = ""
        catre = New List(Of String)()
        Dim parti() As String = linie.Split(New String() {" | "}, 3, StringSplitOptions.None)
        If parti.Length < 3 Then Return False
        If parti(1).Trim() <> "TRIMITERE OK" Then Return False
        Dim mS As String = "Subiect: """
        Dim eC As String = """ | Catre: "
        Dim fF As String = " | Fisiere: "
        Dim d As String = parti(2)
        If Not d.StartsWith(mS) Then Return False
        Dim i2 As Integer = d.IndexOf(eC, mS.Length)
        If i2 < 0 Then Return False
        subiect = d.Substring(mS.Length, i2 - mS.Length)
        Dim i3 As Integer = d.IndexOf(fF, i2 + eC.Length)
        If i3 < 0 Then Return False
        Dim lista As String = d.Substring(i2 + eC.Length, i3 - (i2 + eC.Length))
        For Each a As String In lista.Split(New String() {"; "}, StringSplitOptions.RemoveEmptyEntries)
            If a.Trim().Length > 0 Then catre.Add(a.Trim())
        Next
        Return catre.Count > 0
    End Function

    Private Sub Incarca()
        lst.BeginUpdate()
        lst.Items.Clear()
        Dim celMaiLung As String = ""
        For Each l As String In Linii(Jurnal.Citeste())
            If l.Trim().Length > 0 Then
                lst.Items.Add(l)
                If l.Length > celMaiLung.Length Then celMaiLung = l
            End If
        Next
        If lst.Items.Count = 0 Then lst.Items.Add("(jurnal gol)")
        lst.HorizontalExtent = TextRenderer.MeasureText(celMaiLung, lst.Font).Width + 24
        lst.EndUpdate()
        Dim vizibile As Integer = Math.Max(1, lst.ClientSize.Height \ Math.Max(1, lst.ItemHeight))
        lst.TopIndex = Math.Max(0, lst.Items.Count - vizibile)       ' deruleaza la ultimele intrari
        ActualizeazaSelectie()
    End Sub

    Private Sub ActualizeazaSelectie()
        Dim ok As Boolean = False
        selSubiect = ""
        selCatre = New List(Of String)()
        If lst.SelectedIndex >= 0 Then ok = ParseOk(lst.SelectedItem.ToString(), selSubiect, selCatre)
        btnAnulare.Enabled = ok AndAlso Not trimitere
        If ok Then
            lblSel.Text = "Anulare: """ & PREFIX_STERG & selSubiect & """  catre  " & String.Join("; ", selCatre.ToArray())
        Else
            lblSel.Text = "Selectati o linie TRIMITERE OK pentru a putea anula publicarea."
        End If
    End Sub

    Private Sub LaSelectie(sender As Object, e As EventArgs)
        ActualizeazaSelectie()
    End Sub

    Private Sub LaAfisare(sender As Object, e As EventArgs)
        ' fereastra principala e in dreapta-sus: tinem jurnalul in interiorul ecranului
        Dim wa As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim x As Integer = Math.Max(wa.Left, Math.Min(Me.Left, wa.Right - Me.Width))
        Dim y As Integer = Math.Max(wa.Top, Math.Min(Me.Top, wa.Bottom - Me.Height))
        Me.Location = New Point(x, y)
        Incarca()
    End Sub

    Private Sub LaReimprospateaza(sender As Object, e As EventArgs)
        Incarca()
    End Sub

    Private Sub LaNotepad(sender As Object, e As EventArgs)
        Try
            If Not File.Exists(Jurnal.CaleLog) Then File.WriteAllText(Jurnal.CaleLog, "", New UTF8Encoding(False))
            Diagnostics.Process.Start("notepad.exe", """" & Jurnal.CaleLog & """")
        Catch ex As Exception
            MessageBox.Show(Me, "Nu pot deschide Notepad:" & Environment.NewLine & ex.Message, "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LaGoleste(sender As Object, e As EventArgs)
        If MessageBox.Show(Me, "Stergeti tot continutul jurnalului?", "Jurnal", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Jurnal.Goleste()
            Incarca()
        End If
    End Sub

    Private Sub LaInchide(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' ---------- Anulare publicare ----------
    Private Sub LaAnulare(sender As Object, e As EventArgs)
        If selCatre.Count = 0 Then Return
        Dim subiectNou As String = PREFIX_STERG & selSubiect
        If CONFIRMA_ANULARE Then
            Dim intrebare As String = "Se trimite acum emailul de anulare:" & Environment.NewLine & Environment.NewLine & _
                                      "Subiect: " & subiectNou & Environment.NewLine & _
                                      "Catre: " & String.Join("; ", selCatre.ToArray()) & Environment.NewLine & Environment.NewLine & _
                                      "Continuati?"
            If MessageBox.Show(Me, intrebare, "Anulare publicare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        End If
        Dim j As New SendJob()
        j.S = cfg
        j.Dest = New List(Of String)(selCatre)
        j.Subiect = subiectNou
        j.Mesaj = TEXT_ANULARE
        j.Fisiere = New List(Of String)()
        jobAnulare = j
        Blocheaza(True)
        Dim bw As New BackgroundWorker()
        AddHandler bw.DoWork, AddressOf LucruAnulare
        AddHandler bw.RunWorkerCompleted, AddressOf FinalAnulare
        bw.RunWorkerAsync(j)
    End Sub

    Private Sub LucruAnulare(sender As Object, e As DoWorkEventArgs)
        Mailer.Trimite(DirectCast(e.Argument, SendJob))
    End Sub

    Private Sub FinalAnulare(sender As Object, e As RunWorkerCompletedEventArgs)
        Blocheaza(False)
        If e.Error Is Nothing Then
            Jurnal.Scrie("ANULARE OK", Jurnal.DescrieJob(jobAnulare))
            Incarca()
            MessageBox.Show(Me, "Email de anulare trimis:" & Environment.NewLine & jobAnulare.Subiect, "Anulare publicare", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim t As String = e.Error.GetType().Name & ": " & e.Error.Message
            If e.Error.InnerException IsNot Nothing Then t &= Environment.NewLine & e.Error.InnerException.Message
            Jurnal.Scrie("ANULARE ESUATA", Jurnal.DescrieJob(jobAnulare) & " | Eroare: " & t)
            Incarca()
            MessageBox.Show(Me, "Trimiterea emailului de anulare a esuat:" & Environment.NewLine & t, "Anulare publicare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Blocheaza(blocat As Boolean)
        trimitere = blocat
        For Each c As Control In Me.Controls
            c.Enabled = Not blocat
        Next
        Me.Cursor = If(blocat, Cursors.WaitCursor, Cursors.Default)
        If Not blocat Then ActualizeazaSelectie()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If trimitere Then e.Cancel = True      ' nu se inchide cat timp se trimite
        MyBase.OnFormClosing(e)
    End Sub
End Class
