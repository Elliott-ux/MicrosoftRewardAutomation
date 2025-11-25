using Microsoft.Playwright;
using System.Diagnostics;

namespace MicrosoftRewardAutomation
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Process? edgeProcess = null;
            Source source = new Source();
            try
            {
                KillEdgeProcess();
                // 启动 Edge 浏览器并启用远程调试
                edgeProcess = Process.Start(new ProcessStartInfo
                {
                    FileName= @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    Arguments= "--remote-debugging-port=9222",
                    UseShellExecute=false
                });
                // 等待浏览器启动
                Thread.Sleep(3000);
                // 连接到 Edge 浏览器的调试端口,打开bing网站
                using var pw = await Playwright.CreateAsync();
                await using var browser = await pw.Chromium.ConnectOverCDPAsync("http://127.0.0.1:9222");
                var contexts=browser.Contexts;
                var page = contexts.Count > 0
                    ? (contexts[0].Pages.Count > 0
                        ? contexts[0].Pages[0]
                        : await contexts[0].NewPageAsync())
                    : await (await browser.NewContextAsync()).NewPageAsync();
                await page.GotoAsync("https://www.bing.com");
                for(int i=0;i<source.Poems!.Count;i++)
                {
                    Console.WriteLine($"正在搜索第{i+1}/40条项目");
                    await page.FillAsync("input[id='sb_form_q']", source.Poems[i] );
                    Thread.Sleep(500);
                    await page.PressAsync("input[id='sb_form_q']", "Enter");
                    Console.WriteLine($"第{i+1}条搜索完成");
                    Thread.Sleep(10000);
                }
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("脚本出错辣，下面是出错内容。");
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// 关闭所有Edge进程
        /// </summary>
        static void KillEdgeProcess()
        {
            try 
            {
                var process = new Process();
                process.StartInfo.FileName = "powerShell";
                process.StartInfo.Arguments = "Get-Process msedge -ErrorAction SilentlyContinue | Stop-Process -Force";
                process.StartInfo.UseShellExecute = false;
                process.Start();
                process.WaitForExit(5000);
                if(process.ExitCode==0)
                {
                    Console.WriteLine("已关闭所有Edge进程");
                }
                else
                {
                    Console.WriteLine("关闭Edge进程失败或者无Edge进程");
                }
            }
            catch(Exception e)
            {
                Console.WriteLine($"关闭Edge进程时出错：{e.Message}");
            }
        }
    }

    class Source
    {
        public List<string>? Poems { get; }

        public Source()
        {
            Poems=new List<string>();
            addPoem();
        }

        //添加40首古诗
        private void addPoem()
        {
            Poems?.Add("床前明月光，疑是地上霜。举头望明月，低头思故乡。");
            Poems?.Add("春眠不觉晓，处处闻啼鸟。夜来风雨声，花落知多少。");
            Poems?.Add("红豆生南国，春来发几枝。愿君多采撷，此物最相思。");
            Poems?.Add("月落乌啼霜满天，江枫渔火对愁眠。姑苏城外寒山寺，夜半钟声到客船。");
            Poems?.Add("千山鸟飞绝，万径人踪灭。孤舟蓑笠翁，独钓寒江雪。");
            Poems?.Add("白日依山尽，黄河入海流。欲穷千里目，更上一层楼。");
            Poems?.Add("会当凌绝顶，一览众山小。");
            Poems?.Add("大漠孤烟直，长河落日圆。萧关逢候骑，都护在燕然。");
            Poems?.Add("劝君更尽一杯酒，西出阳关无故人。");
            Poems?.Add("人生自古谁无死，留取丹心照汗青。");
            Poems?.Add("青青子衿，悠悠我心。但为君故，沉吟至今。");
            Poems?.Add("山有木兮木有枝，心悦君兮君不知。");
            Poems?.Add("采菊东篱下，悠然见南山。");
            Poems?.Add("夕阳无限好，只是近黄昏。");
            Poems?.Add("洛阳亲友如相问，一片冰心在玉壶。");
            Poems?.Add("春风又绿江南岸，明月何时照我还。");
            Poems?.Add("人生若只如初见，何事秋风悲画扇。");
            Poems?.Add("无可奈何花落去，似曾相识燕归来。");
            Poems?.Add("天街小雨润如酥，草色遥看近却无。");
            Poems?.Add("竹外桃花三两枝，春江水暖鸭先知。");
            Poems?.Add("独在异乡为异客，每逢佳节倍思亲。");
            Poems?.Add("落红不是无情物，化作春泥更护花。");
            Poems?.Add("千里莺啼绿映红，水村山郭酒旗风。");
            Poems?.Add("夕阳西下，断肠人在天涯。");
            Poems?.Add("山重水复疑无路，柳暗花明又一村。");
            Poems?.Add("劝君莫惜金樽酒，人生得意须尽欢。");
            Poems?.Add("会当击水三千里，立马扬鞭自奋蹄。");
            Poems?.Add("长风破浪会有时，直挂云帆济沧海。");
            Poems?.Add("海内存知己，天涯若比邻。");
            Poems?.Add("莫愁前路无知己，天下谁人不识君。");
            Poems?.Add("天生我材必有用，千金散尽还复来。");
            Poems?.Add("君不见黄河之水天上来，奔流到海不复回。");
            Poems?.Add("抽刀断水水更流，举杯消愁愁更愁。");
            Poems?.Add("长安一片月，万户捣衣声。");
            Poems?.Add("春色满园关不住，一枝红杏出墙来。");
            Poems?.Add("欲把西湖比西子，淡妆浓抹总相宜。");
            Poems?.Add("江南好，风景旧曾谙。日出江花红胜火，春来江水绿如蓝。");
            Poems?.Add("采得百花成蜜后，为谁辛苦为谁甜。");
            Poems?.Add("清明时节雨纷纷，路上行人欲断魂。");
            Poems?.Add("黄河远上白云间，一片孤城万仞山。");
            Poems?.Add("欲穷千里目，更上一层楼。");
        }
    }
}
