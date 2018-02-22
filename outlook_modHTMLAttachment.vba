Option Explicit
Public last_txt As String
Public Const constHead As String = "<html><head><style>" & _
        "body{font-size:pt; font-family:sans-serif}" & _
        ".section{font-size:18pt;text-align: left}" & _
        ".headline{font-size:40pt;font-weight:bold;text-align: center}" & _
        ".author{font-size:18pt;text-align: center}" & _
        ".caption{font-size:16pt;text-align: center}" & _
        ".picture{display: block;margin: auto;width: 40%;}" & _
        "</style><title>"

Function GetArticleLink(olMail As Outlook.PostItem) As String
    Dim els As IHTMLElementCollection, el As HTMLGenericElement
    Dim html As New HTMLDocument
    html.body.innerHTML = olMail.HTMLBody
    Set els = html.getElementsByTagName("a")
    For Each el In els
        If StrComp(Left(el.innerText, Len("view article")), "View article", vbTextCompare) = 0 Then
            GetArticleLink = el.getAttribute("href")
            Exit For
        End If
    Next
End Function

Sub ProcessRssCurrentItem()
    Dim s As Outlook.PostItem
    On Error Resume Next
    For Each s In Application.ActiveExplorer.Selection
        If Err.Number <> 0 Then: Exit Sub
        On Error GoTo 0
        StripAndResave s.entryID
    Next
    On Error GoTo 0
End Sub

Public Sub StripAndResave(entryID As String)
    Dim olMail As PostItem, fn As String, txt As String
    Set olMail = getItem(entryID)                                                           'get olMail item
    fn = getFn                                                                              'get fn to use
    deleteExtraAttachments olMail                                                           'delete all but first attachment
    If saveTempFile(fn, olMail) = 0 Then: Exit Sub                                          'save temp, exit if no attachments
    txt = getText(fn)                                                                       'get file contents
    stripPreviousSubjectMark olMail
    Select Case olMail.Parent
        Case "WashPo National", "WashPo Politics"
            txt = parseHTML_washPo(txt, olMail.Parent, olMail.Subject)                      'do stuff with the raw HTML
        Case "National Review"
            txt = parseHTML_nr(txt, olMail.Parent, olMail.Subject)                          'do stuff with the raw HTML
        Case "NYT Home Page"
            txt = parseHTML_nyt(txt, olMail.Parent, olMail.Subject)                         'do stuff with the raw HTML
        Case "The Atlantic"
            txt = parseHTML_atlantic(txt, olMail.Parent, olMail.Subject, olMail.SenderName) 'do stuff with the raw HTML
        Case "Five Thirty Eight"
            txt = parseHTML_fte(txt, olMail.Parent, olMail.Subject)                       'do stuff with the raw HTML
        Case "Lawfare"
            txt = parseHTML_lf(txt, olMail.Parent, olMail.Subject)                       'do stuff with the raw HTML
        Case Else
            'do nothing
    End Select
    If txt = "Error" Then
        markSubject olMail, "Error Processing"
    ElseIf txt = "Video" Then
        markSubject olMail, "Video"
        deleteMail olMail
    Else                                                                                    'process the mail
 '      deleteAttachment olMail                                                             'delete current attachment
        saveFile fn, txt                                                                    'save txt as file
        setMailBody txt, olMail
 '      attachFile fn, olMail                                                               'attach txt to file
        markSubject olMail, "Processed"
        deleteFile fn                                                                       'delete file
    End If
End Sub

Public Sub deleteMail(olMail As Outlook.PostItem)

    Set olMail = olMail.Move(Session.GetDefaultFolder(olFolderDeletedItems))
    Set olMail = Session.GetItemFromID(olMail.entryID)
    olMail.Delete
End Sub

Public Sub deleteExtraAttachments(olMail As Outlook.PostItem)
    Dim i As Long, a As Outlook.Attachment
    For Each a In olMail.Attachments
        If a.fileName <> "Full article.htm" Then
            a.Delete
        End If
    Next
    For i = 2 To olMail.Attachments.Count
        olMail.Attachments(i).Delete
    Next
    olMail.Save
End Sub

Public Sub setMailBody(txt As String, olMail As Outlook.PostItem)
    txt = "<a href='" & GetArticleLink(olMail) & "'>View Article</a><br />" & txt
    olMail.HTMLBody = txt
    olMail.Save
End Sub

Public Function getItem(entryID As String) As Outlook.PostItem
    Dim ns As Outlook.NameSpace
    Set ns = Outlook.GetNamespace("MAPI")
    Set getItem = ns.GetItemFromID(entryID)                             'item passed in
End Function

Public Function getText(fn As String) As String
    Dim request As New MSXML2.XMLHTTP
    Dim converter As New ADODB.Stream, r As String
    request.Open "GET", fn, False
    request.Send
    converter.Open
    converter.Type = adTypeBinary
    converter.Write request.responseBody
    converter.Position = 0
    converter.Type = adTypeText
    converter.Charset = "utf-8"
    r = converter.ReadText
    getText = r
    converter.Close
    request.abort
End Function

Public Function getFn() As String
    Randomize
    getFn = Environ$("temp") & "\" & Int((1000000 - 1 + 1) * Rnd + 1) & ".html"
End Function

Public Function saveTempFile(fn As String, olMail As Outlook.PostItem) As Integer
    If olMail.Attachments.Count = 0 Then
        olMail.Subject = "No Article - " & olMail.Subject
        olMail.Save
        saveTempFile = 0
    Else
        olMail.Attachments(1).SaveAsFile (fn)
        saveTempFile = 1
    End If
End Function

Public Sub saveFile(fn As String, txt As String)
    Dim fso As New Scripting.FileSystemObject, f As TextStream
    Set f = fso.CreateTextFile(fn, True, True)
    f.Write (txt)
    f.Close
End Sub
Sub markSubject(olMail As Outlook.PostItem, str As String)
    olMail.Subject = str & " - " & StripEm(olMail.Subject)
    olMail.Save
End Sub

Sub stripPreviousSubjectMark(olMail As Outlook.PostItem)
    Dim orig As String
    On Error Resume Next
    orig = Left(olMail.Subject, InStr(1, olMail.Subject, " - ", vbTextCompare) - 1)
    If Err.Number = 0 Then
        Select Case orig
            Case "Processed", "Error", "Video", "Error Processing"
                olMail.Subject = Mid(olMail.Subject, Len(orig) + 4, Len(olMail.Subject))
                olMail.Save
        End Select
    End If
    Err.Clear
    On Error GoTo 0
End Sub
Public Sub attachFile(fn As String, olMail As Outlook.PostItem)
    olMail.Attachments.Add (fn)
    olMail.Save
End Sub

Public Sub deleteFile(fn As String)
    Dim fso As New Scripting.FileSystemObject
    fso.deleteFile (fn)
End Sub

Public Sub deleteAttachment(olMail As Outlook.PostItem)
    On Error Resume Next
    olMail.Attachments(1).Delete
    On Error GoTo 0
    olMail.Save
End Sub
Public Function parseHTML_nr(txt As String, src As String, subj As String) As String
    Dim html As New HTMLDocument
    Dim r As String
    Dim els2 As IHTMLDOMChildrenCollection
    Dim els As IHTMLElementCollection, el As HTMLGenericElement
    html.body.innerHTML = txt
    'head info
    r = r & constHead & StripEm(subj) & "</title></head><body>" & vbCrLf
    'get section
    Set els = html.getElementsByTagName("h1")
    For Each el In els
        If el.className = "blog-nav-title" Then
            r = r & "<div class='section'>" & src & " - " & StripEm(el.innerText) & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get headline
    r = r & "<div class='headline'>" & subj & "</div>" & vbCrLf
    'get reporter
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.getAttribute("itemprop") = "author" Then
            r = r & "<div class='author'>" & el.getAttribute("content") & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get article text
    Set els2 = html.querySelectorAll("p,blockquote, li, ul, img")
    Dim x As Long
    For x = 0 To els2.Length - 1
        Set el = els2(x)
        If StrComp(el.className, "pullquote", vbTextCompare) <> 0 And (StrComp(el.parentElement.getAttribute("itemprop"), "articleBody", vbTextCompare) = 0 Or StrComp(Left(el.className, 4), "CxSp", vbTextCompare) = 0 Or StrComp(el.className, "article-ad-desktop-position", vbTextCompare) = 0 Or StrComp(el.className, "article-ad-mobile-position", vbTextCompare) = 0 Or StrComp(el.parentElement.className, "twitter-tweet", vbTextCompare) = 0) Then
            If StrComp(el.tagName, "blockquote", vbTextCompare) = 0 Or StrComp(el.parentElement.tagName, "blockquote", vbTextCompare) = 0 Or StrComp(el.parentElement.className, "twitter-tweet", vbTextCompare) = 0 Then
                r = r + "<br /><div class='article'><blockquote>" & StripEm(el.innerHTML) & "</blockquote></div>" & vbCrLf
            ElseIf StrComp(el.tagName, "img", vbTextCompare) = 0 Then
                r = r & "<img src='" & el.getAttribute("data-src") & "' />"
            Else
                r = r + "<br /><div class='article'>" & StripEm(el.innerHTML) & "</div>" & vbCrLf
            End If
        End If
    Next
    r = r & "</body></html>"
    parseHTML_nr = r
End Function
Public Function parseHTML_fte(txt As String, src As String, subj As String) As String
    On Error Resume Next
    Dim startPos As Long, endPos As Long, article As String
    txt = StripEm(txt)
    startPos = InStr(txt, "<article")
    endPos = InStr(txt, "</article")
    article = "<title>" & subj & "</title>"
    'article = article & "<h1>" & subj & "</h1>"
    article = article & Mid(txt, startPos, (endPos - startPos) + 10)
    If Err.Number <> 0 Then
        parseHTML_fte = "Error"
    Else
        parseHTML_fte = article
    End If
    On Error GoTo 0
End Function

Public Function parseHTML_atlantic(txt As String, src As String, subj As String, from As String) As String
    Dim html As New HTMLDocument
    Dim r As String, body As String
    Dim els As IHTMLElementCollection, el As HTMLGenericElement, el2 As HTMLGenericElement
    html.body.innerHTML = txt
    'head info
    r = r & constHead & subj & "</title></head><body>" & vbCrLf
    'get section
    r = r & "<div class='section'>The Atlantic</div>" & vbCrLf
    'get headline
    Set els = html.getElementsByTagName("h1")
    For Each el In els
        r = r & "<div class='headline'>" & StripEm(el.innerText) & "</div>" & vbCrLf
        Exit For
    Next
    'get reporter
    r = r & "<div class='author'>" & from & "</div>" & vbCrLf
    'get header image
    Set els = html.getElementsByTagName("img")
    For Each el In els
        If el.parentElement.className = "img" Then
            r = r & "<img class='picture' src='" & el.getAttribute("src") & "' />" & vbCrLf
            Exit For
        End If
    Next
    'get caption
    Set els = html.getElementsByTagName("p")
    For Each el In els
        If el.getAttribute("itemprop") = "description" Then
            r = r & "<div class='caption'>" & StripEm(el.innerText) & "</div>" & vbCrLf
            Exit For
        End If
    Next
    r = r & recurseElements(html.body.innerHTML) & "</body></html>"
    parseHTML_atlantic = r
End Function

Public Function parseHTML_nyt(txt As String, src As String, subj As String) As String
    Dim html As New HTMLDocument
    Dim r As String
    Dim els As IHTMLElementCollection, el As HTMLGenericElement
    html.body.innerHTML = txt
    
    'see if video
    Set els = html.getElementsByTagName("a")
    For Each el In els
        If el.getAttribute("title") = "TimesVideo" Then
            parseHTML_nyt = "Video"
            Exit Function
         End If
    Next

    'head info
    r = r & constHead & StripEm(subj) & "</title></head><body>" & vbCrLf
    'get section
    Set els = html.getElementsByClassName("kicker-label")
    For Each el In els
        r = r & "<div class='section'>" & src & " - " & StripEm(el.innerText) & "</div>" & vbCrLf
        Exit For
    Next
    'get headline
    Set els = html.getElementsByTagName("h1")
    For Each el In els
        r = r & "<div class='headline'>" & StripEm(el.innerText) & "</div>" & vbCrLf
        Exit For
    Next
    'get reporter
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.className = "byline-author" Then
            r = r & "<div class='author'>" & StripEm(el.innerText) & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get header image
    Set els = html.getElementsByTagName("img")
    For Each el In els
        If el.className = "media-viewer-candidate" Then
            r = r & "<img class='picture' src='" & el.getAttribute("src") & "' />" & vbCrLf
            Exit For
        End If
    Next
    'get caption
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.className = "caption-text" Or el.className Like "*captionText*" Then
            r = r & "<div class='caption'>" & StripEm(el.innerText) & "</div>" & vbCrLf
            Exit For
        End If
    Next
    r = r & recurseElements(html.body.innerHTML) & "</body></html>"
    parseHTML_nyt = r
End Function


Public Function parseHTML_washPo(txt As String, src As String, subj As String) As String
    Dim html As New HTMLDocument
    Dim r As String
    Dim els As IHTMLElementCollection, el As HTMLGenericElement
    html.body.innerHTML = txt
    
    'see if video
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.getAttribute("className") = "page-title" Then
            On Error Resume Next
            If el.PreviousSibling.getAttribute("className") = "nav-subsection" Then
                parseHTML_washPo = "Video"
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
    Next
    'head info
    r = r & constHead & StripEm(subj) & "</title></head><body>" & vbCrLf
    'get section
    Set els = html.getElementsByClassName("headline-kicker")
    For Each el In els
        r = r & "<div class='section'>" & src & " - " & StripEm(el.innerText) & "</div>" & vbCrLf
        Exit For
    Next
    'get headline
    Set els = html.getElementsByTagName("h1")
    For Each el In els
        r = r & "<div class='headline'>" & StripEm(el.innerText) & "</div>" & vbCrLf
        Exit For
    Next
    'get reporter
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.className = "pb-byline" Then
            r = r & "<div class='author'>" & el.innerText & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get header image
    Set els = html.getElementsByTagName("img")
    For Each el In els
        If StrComp(el.className, "unprocessed placeholder", vbTextCompare) = 0 Or StrComp(el.className, "hi-res-lazy courtesy-of-the-lazy-loader", vbTextCompare) = 0 Then
            r = r & "<img class='picture' src='" & el.getAttribute("data-raw-src") & "' />" & vbCrLf
            Exit For
        End If
    Next
    'get caption
    Set els = html.getElementsByTagName("span")
    For Each el In els
        If el.className = "pb-caption" Then
            r = r & "<div class='caption'>" & el.innerText & "</div>" & vbCrLf
            Exit For
        End If
    Next
    r = r & recurseElements(html.body.innerHTML) & "</body></html>"
    parseHTML_washPo = r
End Function
Function recurseElements(txt As String) As String
    
    Dim i As Long, a As String
    Dim html As New HTMLDocument
    Dim el As HTMLGenericElement, els As IHTMLDOMChildrenCollection
    Dim stripped As String
    
    html.body.innerHTML = txt
    Set els = html.querySelectorAll("p,blockquote,ul,li,h1,h2,h3,h4")
    
    For i = 0 To els.Length - 1
        Set el = els.Item(i)
        If (el.className Like "*Paragraph-paragraph*" Or el.className = "dropcap" Or el.className = "p1" Or el.className = "field-item" Or el.className = "p2" Or el.className Like "*hed*" Or el.className Like "*story-subheading*" Or el.className = "pg-bodyCopy" Or el.className = "" Or el.className = "story-body-text story-content" Or el.className = "citation" Or el.className = "twitter-tweet") And (el.parentElement.className Like "*field-item*" Or el.parentElement.className = "" Or el.parentElement.className Like "*inline-content*" Or el.parentElement.className Like "*article-body*" Or el.parentElement.className Like "*story-body*" Or el.parentElement.className Like "*content*" Or el.parentElement.className Like "*hed*") Then
            stripped = StripEm(el.innerHTML)
            If last_txt <> el.innerText Then
                If el.tagName = "BLOCKQUOTE" Then
                    a = a & "<br /><div class='article'><blockquote>" & stripped & "</blockquote></div>" & vbCrLf
                Else
                    a = a & "<br /><div class='article'>" & stripped & "</div>" & vbCrLf
                End If
            End If
            last_txt = el.innerText
            a = a & recurseElements(el.innerHTML)
        End If
    Next
    recurseElements = a
End Function


Function StripEm(strSource As String) As String
    StripEm = strSource
End Function



Public Function parseHTML_lf(txt As String, src As String, subj As String) As String
    Dim html As New HTMLDocument
    Dim r As String
    Dim els2 As IHTMLDOMChildrenCollection
    Dim els As IHTMLElementCollection, el As HTMLGenericElement
    html.body.innerHTML = txt
    'head info
    r = r & constHead & StripEm(subj) & "</title></head><body>" & vbCrLf
    'get section
    Set els = html.getElementsByTagName("h6")
    For Each el In els
        If el.className = "article__featured-term" Then
            r = r & "<div class='section'>" & src & " - " & StripEm(el.innerText) & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get headline
    r = r & "<div class='headline'>" & subj & "</div>" & vbCrLf
    'get reporter
    Set els = html.getElementsByTagName("div")
    For Each el In els
        If el.className = "article-top__contributors" Then
            r = r & "<div class='author'>" & el.innerText & "</div>" & vbCrLf
            Exit For
        End If
    Next
    'get article text
    r = r & recurseElements(html.body.innerHTML) & "</body></html>"
    parseHTML_lf = r
End Function
