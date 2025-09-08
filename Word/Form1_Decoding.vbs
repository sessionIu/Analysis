' Form1_Decoding
Public Function b64Decode(EncodedText) As Byte()
    Dim data() As Byte, encodedData() As Byte
    Dim DataLength As Long, EncodedLength As Long
    Dim l As Long, Index As Long
    
    Const Mask1 As Byte = 3
    Const Mask2 As Byte = 15
    Const Shift2 As Byte = 4   
    Const Shift4 As Byte = 16   
    Const Shift6 As Byte = 64    
    
    Dim Base64Lookup() As Byte, Base64Reverse() As Byte
    ReDim Base64Reverse(255) [ 배열 지정 ]
    Base64Lookup = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode) ` A-Z a-z 0-9 + / 문자 저장
    
    
    For l = 0 To 63
        Base64Reverse(Base64Lookup(l)) = l                                                       ' 디코딩 변환 속도를 위한 인덱싱
    Next l
    
   
    encodedData = StrConv(Replace$(Replace$(EncodedText, vbCrLf, ""), "=", ""), vbFromUnicode)   `  CRLF 를 "" 로 , = 를 ""로 변환 후 문자열을 바이트 배열로
    EncodedLength = UBound(encodedData) + 1                                                      `  UBound(배열) 최대 인덱스반환 - VBA 배열은 0부터 시작하므로 + 1
                                                                                                 `  Base64 처리를 위한 vbFromUnicode 변환 [ vbs 유니코드 2바이트 ] [ ANSI 1바이트 ]
    DataLength = (EncodedLength \ 4) * 3                                                         `  Base64 4문자 -> 3 바이트 변환 [ 원본 데이터 3바이트 -> 4문자 인코딩 역순 ]
    Dim m As Long: m = EncodedLength Mod 4                                                       `  4로 나눈 나머지
    If m = 2 Then DataLength = DataLength + 1                                                    `  나머지 2 -> 1바이트 추가 [  base64 는 6비트씩 읽기때문에 ]
    If m = 3 Then DataLength = DataLength + 2                                                    `  나머지 3 -> 2바이트 추가
    ReDim data(DataLength - 1)                                                                   `  실제 출력 배열 ( 0 부터 ~ 시작이므로 -1로 지정 )
    
    For l = 0 To UBound(encodedData) - m Step 4                                                  `  4개씩 나눠서 처리 , m 은 4로 나눴을때의 반복문 종료변수 -> step 통한 배열이동
        Dim ed0 As Long: ed0 = Base64Reverse(encodedData(l))                                     `  첫번째 ~ 4번째 문자열
        Dim ed1 As Long: ed1 = Base64Reverse(encodedData(l + 1))
        Dim ed2 As Long: ed2 = Base64Reverse(encodedData(l + 2))
        Dim ed3 As Long: ed3 = Base64Reverse(encodedData(l + 3))
        
        data(Index) = (ed0 * Shift2) Or (ed1 \ Shift4)                                           `  data 는 ed0 2비트 왼쪽 ,  ed1 4비트 오른쪽 채우기 [0으로]
        data(Index + 1) = ((ed1 And Mask2) * Shift4) Or (ed2 \ Shift2)                           `  위에서 정의한 15 = 이진수 1111 , 00001111 and 로 하위 4비트만 통과 후 왼쪽에 0000 채운 뒤 ed2 에 값2개 버린값과 결합
        data(Index + 2) = ((ed2 And Mask1) * Shift6) Or ed3                                      `  흐름은 같다
        Index = Index + 3
    Next l
    
    Select Case m
        Case 2:
            ed0 = Base64Reverse(encodedData(l))
            ed1 = Base64Reverse(encodedData(l + 1))
            data(Index) = (ed0 * Shift2) Or (ed1 \ Shift4)
        Case 3:
            ed0 = Base64Reverse(encodedData(l))
            ed1 = Base64Reverse(encodedData(l + 1))
            ed2 = Base64Reverse(encodedData(l + 2))
            data(Index) = (ed0 * Shift2) Or (ed1 \ Shift4)
            data(Index + 1) = ((ed1 And Mask2) * Shift4) Or (ed2 \ Shift2)
    End Select
    
    b64Decode = data                                                                                ` 리턴 값
End Function

