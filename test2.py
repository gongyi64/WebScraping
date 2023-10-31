var options = new ChromeOptions
{
    DebuggerAddress = "127.0.0.1:9222"
};

using (var driver = new ChromeDriver(options))
{
    // 任意のブラウザ操作処理 ↓↓↓
    var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
    driver.Url = "https://www.google.com";
    var q = driver.FindElement(By.Name("q"));
    q.SendKeys("Chromium");
    q.Submit();

    wait.Until(ExpectedConditions.TitleIs("Chromium - Google 検索"));
    ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile($"{DateTime.Now.ToString("yyyyMMddHHmmss")}.png");
    // 任意のブラウザ操作処理 ↑↑↑

}

#これは、C＃での記述