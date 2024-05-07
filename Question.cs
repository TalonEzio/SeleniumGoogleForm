namespace SeleniumGoogleForm
{
    /// <summary>
    /// Theo format của NineQuiz.com
    /// </summary>
    internal class Question
    {
        public string Title { get; set; } = string.Empty;
        public string Answer { get; set; } = string.Empty;
        public string AnswerA { get; set; } = string.Empty;
        public string AnswerB { get; set; } = string.Empty;
        public string AnswerC { get; set; } = string.Empty;
        public string AnswerD { get; set; } = string.Empty;
    }
}
