using GemBox.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text;
using static GemBox.Spreadsheet.SpreadsheetInfo;

namespace SeleniumGoogleForm
{
    internal class Program
    {
        
        private static async Task Main()
        {
            Console.InputEncoding = Console.OutputEncoding = Encoding.Unicode;
            var urlList = await File.ReadAllLinesAsync("url.txt");
            var i = 1;

            var bigQuestions = new List<Question>();
            foreach (var url in urlList)
            {
                var fileName = $"Bài trắc nghiêm số {i++}.xlsx";

                try
                {
                    var questions = GetQuestion(url).ToList();
                    bigQuestions.AddRange(questions);
                    await ExportToExcel(questions, fileName);
                    Console.WriteLine($"Export xong {fileName}");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"{e.Message} - {fileName}");
                }
            }

            var chunkSize = 100;
            var questionChunks = bigQuestions
                .Select((question, index) => new { Question = question, Index = index })
                .GroupBy(x => x.Index / chunkSize)
                .Select(g => g.Select(x => x.Question).ToList())
                .ToList();

            for (int item = 0; item < questionChunks.Count; item++)
            {
                var fileIndex = item + 1;
                var fileName = $"Bài trắc nghiệm TACN_Split-{chunkSize}_{fileIndex}.xlsx";
                await ExportToExcel(questionChunks[item], fileName);
                Console.WriteLine($"Xuất xong {fileName}");
            }

            Console.ReadLine();
        }
        private static async Task ExportToExcel(IEnumerable<Question> questions, string fileName)
        {
            SetLicense("TalonEzio-Cracked-Hehe");

            var workbook = new ExcelFile();

            ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet1");

            worksheet.Cells["A1"].Value = "Câu hỏi";
            worksheet.Cells["B1"].Value = "Đáp án";
            worksheet.Cells["C1"].Value = "Câu trả lời A";
            worksheet.Cells["D1"].Value = "Câu trả lời B";
            worksheet.Cells["E1"].Value = "Câu trả lời C";
            worksheet.Cells["F1"].Value = "Câu trả lời D";

            int row = 1;
            foreach (var question in questions)
            {
                worksheet.Cells[row, 0].Value = question.Title;
                worksheet.Cells[row, 1].Value = question.Answer;
                worksheet.Cells[row, 2].Value = question.AnswerA;
                worksheet.Cells[row, 3].Value = question.AnswerB;
                worksheet.Cells[row, 4].Value = question.AnswerC;
                worksheet.Cells[row, 5].Value = question.AnswerD;
                row++;
            }

            var directory = Path.Combine(Environment.CurrentDirectory, "Exports");
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            await using var stream = File.Create(Path.Combine(directory, fileName));
            workbook.Save(stream, SaveOptions.XlsxDefault);
        }

        static IEnumerable<Question> GetQuestion(string url)
        {
            string[] indexAnswers = ["A", "B", "C", "D"];
            var options = new ChromeOptions();
            options.AddArgument("--headless");//disable open chrome
            IWebDriver driver = new ChromeDriver(options);

            driver.Navigate().GoToUrl(url);

            IReadOnlyList<IWebElement> elements = driver.FindElements(By.CssSelector(".Qr7Oae"));
            var questionList = new List<Question>();

            foreach (var element in elements)
            {
                var questionElement = element.FindElement(By.CssSelector(".cTDvob.D1wxyf.RjsPE"));

                var questionContent = questionElement.GetAttribute("textContent");

                questionContent = questionContent
                    .Replace("\n", " ")
                    .Replace("\r", " ")
                    .Replace("\t", " ")
                    .Replace("*", "");

                string[] ignoreList = ["Email", "*", "Địa chỉ Email:", "Lớp:", "Họ và tên"];

                if (ignoreList.Contains(questionContent)) continue;

                var answersElements = element.FindElements(By.CssSelector(".aDTYNe.snByac.kTYmRb.OIC90c"));

                var answerContents = answersElements.Select(x => x.Text).ToList();
                if(!answerContents.Any() ) continue;

                var question = new Question()
                {
                    Title = questionContent,
                    AnswerA = answerContents[0],
                    AnswerB = answerContents[1],
                    AnswerC = answerContents[2],
                    AnswerD = answerContents[3],
                    Answer = answerContents.Count == 5 ? answerContents[4] : ""
                };

                question.Answer = question.Answer switch
                {
                    var answer when answer == question.AnswerA => "A",
                    var answer when answer == question.AnswerB => "B",
                    var answer when answer == question.AnswerC => "C",
                    var answer when answer == question.AnswerD => "D",
                    _ => question.Answer
                };

                //Tìm Kết quả đúng
                if (string.IsNullOrEmpty(question.Answer))
                {
                    var answers = element.FindElements(By.ClassName("yUJIWb"));

                    for (var i = 0; i < answers.Count; ++i)
                    {
                        var answer = answers[i];
                        var answerInnerHtml = answer.GetAttribute("innerHTML");

                        if (!answerInnerHtml.Contains("Chính xác")) continue;

                        question.Answer = indexAnswers[i];
                        break;
                    }
                }
                questionList.Add(question);

            }
            driver.Quit();
            return questionList;
        }

    }
}
