Function doFirebaseSignIn()

    Dim objHTTP, jsonResult As Object
    Dim google_signin_url As String
    Dim jsonRequestBody, signinResponse As String
    Dim fb_email, fb_pwd, fb_apikey As String
    Dim fb_auth_token As String
    
    'Credentials
    fb_email = "<e-mail address>"
    fb_pwd = "<password">
    fb_apikey = "<project-app API key>"
    fb_auth_token = ""
    google_signin_url = "https://www.googleapis.com/identitytoolkit/v3/relyingparty/verifyPassword?key=" & fb_apikey
    
    If fb_email <> "" And fb_pwd <> "" Then
    
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        
        jsonRequestBody = "{""email"":""" & fb_email & """,""password"":""" & fb_pwd & """,""returnSecureToken"": true}"
        
        objHTTP.Open "POST", google_signin_url, False
        objHTTP.setRequestHeader "Content-type", "application/json"
        objHTTP.Send (jsonRequestBody)
        
        signinResponse = objHTTP.responseText
        'idToken is the api key for DB access
        
        Set jsonResult = JsonConverter.ParseJson(signinResponse)
        fb_auth_token = jsonResult("idToken")
        
        Set jsonResult = Nothing
        Set objHTTP = Nothing
        
    End If
    
    doFirebaseSignIn = fb_auth_token
    
End Function