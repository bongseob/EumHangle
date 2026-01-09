## 아래아 한글 사용 시 보안 모듈 적용 하기 ##
한글 파일을 Automation 작업하기 위해 파일에 접근하거나 이미지를 포함할 때 "보안 모듈 허용" 팝업이 뜬다.
자동화 과정에서 팝업이 뜨면서 멈추기 때문에 필수적으로 해결이 되어야 한다.
- 나의 해결 방안
1. 먼저 다음과 같은 메서드를 만들어 둔다.
```cs
        private static void RegisterSecurityModule()
        {
            try
            {
                string moduleFileName = "FilePathCheckerModuleExample.dll";
                string executablePath = Application.StartupPath + @"\" + moduleFileName;

                // 레지스트리 키 경로
                string keyPath = @"Software\HNC\HwpAutomation\Modules";

                if (!System.IO.File.Exists(executablePath))
                {
                    MessageBox.Show("보안 모듈 파일이 존재하지 않습니다: " + executablePath, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("FilePathCheckerModuleExample", executablePath, RegistryValueKind.String);
                        //key.SetValue("FilePathCheckDLL", executablePath, RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                // 레지스트리 접근 실패 시 오류 메시지를 보여주지만, 앱 실행은 계속됩니다.
                MessageBox.Show($"한/글 보안 모듈 레지스트리 등록 실패: {ex.Message}", "레지스트리 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
```
2. 이 메서드는 Form_Load 에서 호출했다.
```cs
        private void Form1_Load(object sender, EventArgs e)
        {
            // 한글 파일 열 때 경고문 출현 방지 (레지스트리에 보안모듈 추가)
            RegisterSecurityModule();
        }
```
3. 그리고 이미지 추가의 경우는 다음처럼 처리해 주었다.
   이미지를 추가할 때 _hwp.RegisterModule 을 호출했다. 호출에 사용된 인자값은 레지스트리에 등록된 값과 일치시켜줘야 한다.
```cs
        public void InsertImage(string imagePath)
        {
            RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");

            bool embedded = true;    // 이미지 파일을 문서내에 포함할지 여부 (True/False). 생략하면 true
            int sizeoption = 2;      // 삽입할 그림의 크기 옵션 0: 원본 크기 1: 사용자가 지정한 크기 (mmPicWidth, mmPicHeight 값 사용) 2:문서에 맞게 자동 조절
            bool reverse = false;    // 이미지 반전 유무
            bool watermark = false;  // 워터마크 여부
            int effect = 0;          // 이미지 효과 0: 없음
            int mmPicWidth = 0;      // 그림 가로 크기 (sizeoption이 1일 때 사용, 단위: mm)
            int mmPicHeight = 0;     // 그림 세로 크기 (sizeoption이 1일 때 사용, 단위: mm)
            _hwp.InsertPicture(imagePath, embedded, sizeoption, reverse, watermark, effect, mmPicWidth, mmPicHeight);
        }

        public void RegisterModule(string dllName, string moduleName)
        {
            _hwp.RegisterModule(dllName, moduleName);
        }
```
RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample"); 
위 함수의 인자값은 자신의 레지스트리 등록 내용과 일치 시켜야 한다.
<img width="1312" height="342" alt="image" src="https://github.com/user-attachments/assets/c1539ede-5c3a-4211-8a4a-7105ca91672a" />

## 표를 채울 때 주의 사항 ##
표에서 셀합치기가 적용된 경우 행과 열번호가 변경이 생긴다.
행과 열번호는 가장 위쪽과 가장 좌측을 기준으로 번호가 적용되기 때문에 첫 행이나 열에 셀합치기가 적용된 경우 반복문으로 처리할 때 주의를 해야 한다.
그래서 나는 기준이 되는 곳만 행, 열의 번호로 적용하고 이후 컬럼은 MoveDownCell 액션으로 제어했다.
