using Microsoft.Playwright;

namespace MicrosoftRewardAutomation
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using var pw = await Playwright.CreateAsync();
                await using var browser = await pw.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Channel="msedge",
                    Headless=false
                });
                var page = await browser.NewPageAsync();
                await page.GotoAsync("https://www.bing.com");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("脚本出错辣，下面是出错内容。");
                Console.WriteLine(e.Message);
            }
        }
    }
}
