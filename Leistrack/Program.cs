using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;

namespace Leistrack
{
    class Program
    {
        const string VERSION = "V1.0";
        const string APPNAME = "Leistrack";
        static ProgramDataMeta metaData;
        static ProgramDataForYear programDataForYear;
        static LeistungDate currentDate;
        static string dataPath = "";

        static void Main(string[] args)
        {
            RunLeistrack();
        }

        static void RunLeistrack()
        {
            OnStartup();
            ProgramLoop();            
            OnShutdown();
        }

        static void WelcomeUser()
        {
            ClearConsole();
            Console.Write("Welcome to "); WriteColored($"{APPNAME} ", ConsoleColor.Yellow); WriteColored(VERSION, ConsoleColor.Green); Console.WriteLine("!");
            Console.WriteLine($"Current date set to {LeistungDateToRealDateString(currentDate)}");
            Console.WriteLine("You can start by typing \"help\"");
        }

        static string LeistungDateToRealDateString(LeistungDate date)
        {
            DateTime today = programDataForYear.startDate + TimeSpan.FromDays(date.daysSinceFirstSchoolDay);
            return $"{today.Day}/{today.Month}/{today.Year}";
        }

        static int WeekStartToDaysSinceSchoolStarted(int week)
        {
            if (week == 0)
            {
                return 0;
            }
            int daysSinceSchoolStarted = week * 7;
            int schoolStartWeekday = (int)programDataForYear.startDate.DayOfWeek;
            return daysSinceSchoolStarted - schoolStartWeekday+1;
        }

        static int DaysSinceSchoolStartedToWeek(int daysSinceSchoolStarted)
        {
            return (daysSinceSchoolStarted + (int)programDataForYear.startDate.DayOfWeek) / 7;
        }

        static int DateStringToDaysSinceSchoolStarted(string date)
        {
            (bool isValid, DateTime returnDate) = DateStringToDateTime(date);
            int daysSinceSchoolStarted = (int)(returnDate - programDataForYear.startDate).TotalDays;

            if (!isValid || daysSinceSchoolStarted < 0)
            {
                return -1;
            }
            return daysSinceSchoolStarted;
        }

        static (bool isInputValid, DateTime date) DateStringToDateTime(string date)
        {
            int year = 0;
            int month = 0;
            int day = 0;
            string[] dateElements = date.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (dateElements.Length != 3 ||
                !int.TryParse(dateElements[0], out day) ||
                !int.TryParse(dateElements[1], out month) ||
                !int.TryParse(dateElements[2], out year))
            {
                return (false, DateTime.Now);
            }

            DateTime validDateTest;
            try
            {
                validDateTest = new DateTime(year, month, day);
            }
            catch (ArgumentOutOfRangeException e)
            {
                return (false, DateTime.Now);
            }
            return (true, validDateTest);
        }

        static void ProgramLoop()
        {
            Console.WriteLine();
            string input = Console.ReadLine();
            ClearConsole();

            while (input != "quit")
            {
                if (!ReadCommand(input))
                {
                    Console.WriteLine("Invalid input. Type \"help\" for available commands.");
                }
                Console.WriteLine();
                input = Console.ReadLine();
                ClearConsole();
            }
        }

        static void OnShutdown()
        {
            SaveDialog();
            Console.WriteLine("Press any key to close the app.");
            Console.ReadKey();
        }

        static void SaveDialog()
        {
            Console.WriteLine("Save all changes?");

            string input = Console.ReadLine();
            while (input != "yes" && input != "no")
            {
                Console.WriteLine("Save all changes?");
                input = Console.ReadLine();
            }
            if (input == "yes")
            {
                SaveMetaData();
            }
        }

        static bool ReadCommand(string input)
        {
            string[] commandElements = input.Split(" ", StringSplitOptions.RemoveEmptyEntries);

            if (commandElements.Length == 0)
            {
                return false;
            }

            if (programDataForYear.subjects.Contains(commandElements[0]))
            {
                ReadSubjectCommand(commandElements);
                return true;
            }

            switch (commandElements[0])
            {
                case "info":
                    ReadInfoCommand(commandElements);
                    return true;
                case "setdate":
                    ReadSetdateCommand(commandElements);
                    return true;
                case "period":
                    ReadPeriodCommand(commandElements);
                    return true;
                case "subjects":
                    PrintSubjects();
                    return true;
                case "help":
                    PrintHelp();
                    return true;
                default:
                    return false;
            }
        }

        static void PrintSubjects()
        {
            Console.WriteLine("All existing subjects for the current grade are: " + string.Join(", ", programDataForYear.subjects));
        }

        static void PrintHelp()
        {
            Console.WriteLine("1)\"info\" is used to show the time spent on subjects for the current day.");
            Console.WriteLine("2)\"setdate\" is used to change the current day (useful for logging progress which you've forgotten).");
            Console.WriteLine("3)\"period\" is used to show the time spent on subjects for a given period.");
            Console.WriteLine("4)\"subjects\" is used to show all subjects for a given grade.");
            Console.WriteLine("5)\"(subject)\" is used to track your time spent on a subject.");
            Console.WriteLine("For a more detailed description on any given command and its usage, type the command.");
        }

        static void ReadPeriodCommand(string[] commandElements)
        {
            if (commandElements.Length > 1)
            {

                switch (commandElements[1])
                {
                    case "week":
                        ReadInfoForGivenWeekCommand(commandElements);
                        return;
                    default:
                        ReadInfoForGivenPeriodCommand(commandElements);
                        return;
                }


            }
            Console.WriteLine("You can check your total time spent on all subjects for a given period using: \"period (start date) (end date)\"");
            Console.WriteLine($"for example type \"period {LeistungDateToRealDateString(programDataForYear.leistungDates[0])} {LeistungDateToRealDateString(programDataForYear.leistungDates[GetDaysFromStartToNow()])}\"");
            Console.WriteLine("You can check your total time spent on all subjects for a given week using: \"period (week since school start)\" 0 is the first week.");

        }

        static void ReadInfoForGivenWeekCommand(string[] commandElements)
        {
            if (commandElements.Length == 3)
            {
                switch (commandElements[2])
                {
                    case "this":
                        PrintInfoForCurrentWeek();
                        return;
                    default:
                        int week = 0;
                        if(!int.TryParse(commandElements[2], out week) || week < 0)
                        {
                            break;
                        }
                        PrintInfoForGivenWeek(week);
                        return;
                }
                
            }
            Console.WriteLine("Invalid input! Make sure the week is a whole non-negative number or use \"this\" for the current week.");
        }

        static void ReadInfoForGivenPeriodCommand(string[] commandElements)
        {
            if (commandElements.Length == 3)
            {
                int periodStartDaysSinceSchoolStartedCheck = DateStringToDaysSinceSchoolStarted(commandElements[1]);
                int periodEndDaysSinceSchoolStartedCheck = DateStringToDaysSinceSchoolStarted(commandElements[2]);

                if (periodEndDaysSinceSchoolStartedCheck != -1 &&
                    periodStartDaysSinceSchoolStartedCheck != -1 &&
                    periodEndDaysSinceSchoolStartedCheck - periodStartDaysSinceSchoolStartedCheck > 0)
                {
                    PrintInfoForGivenPeriod(periodStartDaysSinceSchoolStartedCheck, periodEndDaysSinceSchoolStartedCheck);
                    return;
                }
            }
            Console.WriteLine("Invalid input! Make sure the dates are correct and are formated like this: dd/mm/yyyy");
            return;
        }

        static void PrintInfoForGivenWeek(int week)
        {
            int weekStart = WeekStartToDaysSinceSchoolStarted(week);
            int weekEnd = WeekStartToDaysSinceSchoolStarted(week + 1) - 1;
            PrintInfoForGivenPeriod(weekStart, weekEnd);
            return;
        }

        static void PrintInfoForCurrentWeek()
        {
            int week = DaysSinceSchoolStartedToWeek(currentDate.daysSinceFirstSchoolDay);
            int weekStart = WeekStartToDaysSinceSchoolStarted(week);
            int weekEnd = WeekStartToDaysSinceSchoolStarted(week + 1) - 1;
            PrintInfoForGivenPeriod(weekStart, weekEnd);
            return;
        }

        static void PrintInfoForGivenPeriod(int startInDaysSinceSchoolStarted, int endInDaysSinceSchoolStarted)
        {
            int[] timeInMinutesForPeriod = new int[programDataForYear.subjects.Count];
            int[] realTimeInMinutesForPeriod = new int[programDataForYear.subjects.Count];

            if (programDataForYear.leistungDates.Count < endInDaysSinceSchoolStarted + 1)
            {
                CreateEmptyDaysTo(endInDaysSinceSchoolStarted + 1);
            }

            for (int i = startInDaysSinceSchoolStarted; i < endInDaysSinceSchoolStarted + 1; i++)
            {
                for (int j = 0; j < programDataForYear.subjects.Count; j++)
                {
                    timeInMinutesForPeriod[j] += programDataForYear.leistungDates[i][programDataForYear.subjects[j]].minutes;
                    realTimeInMinutesForPeriod[j] += programDataForYear.leistungDates[i][programDataForYear.subjects[j]].realMinutes;
                }
            }

            Console.WriteLine("{0} - {1}", LeistungDateToRealDateString(programDataForYear.leistungDates[startInDaysSinceSchoolStarted]), LeistungDateToRealDateString(programDataForYear.leistungDates[endInDaysSinceSchoolStarted]));
            for (int i = 0; i < programDataForYear.subjects.Count; i++)
            {
                Console.WriteLine("{0,-8}: time: {1,9}, real time: {2,9}", programDataForYear.subjects[i], ConvertMinutesToHoursString(timeInMinutesForPeriod[i]), ConvertMinutesToHoursString(realTimeInMinutesForPeriod[i]));
            }
        }

        static void ReadSetdateCommand(string[] commandElements)
        {
            int daysSinceSchoolStarted = 0;
            if (commandElements.Length == 2)
            {
                daysSinceSchoolStarted = DateStringToDaysSinceSchoolStarted(commandElements[1]);
                if (daysSinceSchoolStarted != -1)
                {
                    SetCurrentDay(daysSinceSchoolStarted);
                    ClearConsole();
                    Console.WriteLine($"Current date set to {LeistungDateToRealDateString(currentDate)}");
                    return;
                }
            }
            else if (commandElements.Length == 1)
            {
                DateTime schoolStart = programDataForYear.startDate;
                DateTime tomorrow = DateTime.Now + TimeSpan.FromDays(1);
                Console.WriteLine($"You can change the current day by typing \"setdate (date)\"");
                Console.WriteLine($"For example type \"setdate {tomorrow.Day}/{tomorrow.Month}/{tomorrow.Year}\" to change the date to tomorrow");
                Console.WriteLine($"    * the earliest possible date is {schoolStart.Day}/{schoolStart.Month}/{schoolStart.Year}");
                return;
            }
            Console.WriteLine("Invalid input! Type \"setdate\" for help.");
        }

        static void ClearConsole(bool showDate = true)
        {
            Console.Clear();
            PrintWatermark(showDate);
        }

        static void PrintWatermark(bool showDate = true)
        {
            if (showDate)
            {
                WriteColoredLine($"{APPNAME} {LeistungDateToRealDateString(currentDate)}", ConsoleColor.Yellow);
            }
            else
            {
                WriteColoredLine($"{APPNAME}", ConsoleColor.Yellow);
            }
            Console.WriteLine();
        }

        static void ReadInfoCommand(string[] commandElements)
        {
            if (commandElements.Length < 2)
            {
                PrintInfoForGivenLeistungDate(currentDate);
            }
        }

        static void PrintInfoForGivenLeistungDate(LeistungDate date)
        {
            for (int i = 0; i < programDataForYear.subjects.Count; i++)
            {
                Console.WriteLine("{0,-8}: time: {1,9}, real time: {2,9}", programDataForYear.subjects[i], ConvertMinutesToHoursString(date[programDataForYear.subjects[i]].minutes), ConvertMinutesToHoursString(date[programDataForYear.subjects[i]].realMinutes));
            }
        }

        static void ReadSubjectCommand(string[] commandElements)
        {
            if (commandElements.Length > 1)
            {
                switch (commandElements[1])
                {
                    case "time":
                        ReadSubjectAnyTimeCommand(commandElements, false);
                        return;
                    case "realtime":
                        ReadSubjectAnyTimeCommand(commandElements, true);
                        return;
                }
            }
            Console.WriteLine("You can:\n" +
                             $" 1) check your time spent on the subject by typing \"{commandElements[0]} time\"\n" +
                             $" 2) check your real time spent on the subject by typing \"{commandElements[0]} realtime\"\n" +
                             $" 3) change your time spent on the subject by typing \"{commandElements[0]} time (minutes)\"\n" +
                             $"    for example type \"{commandElements[0]} time +60\" to add 60 minutes or \"{commandElements[0]} time -60\" to subtract 60 minutes.\n" +
                             $" 4) change your real time spent on the subject by typing \"{commandElements[0]} realtime (minutes)\"\n" +
                             $" 5) set your time spent on the subject by typing \"{commandElements[0]} time set (minutes)\" to set the time spent on a subject to a fixed amount\n" +
                             $" 6) set your real time spent on the subject by typing \"{commandElements[0]} realtime set (minutes)\"");
        }

        static void ReadSubjectAnyTimeCommand(string[] commandElements, bool isRealTime)
        {
            if (commandElements.Length > 2)
            {
                switch (commandElements[2])
                {
                    case "set":
                        int toMinutes = 0;
                        if (commandElements.Length > 3 && int.TryParse(commandElements[3], out toMinutes))
                        {
                            if (isRealTime)
                            {
                                currentDate[commandElements[0]] = (currentDate[commandElements[0]].minutes, toMinutes);
                                Console.WriteLine($"Real time in {commandElements[0]} set to: {toMinutes} minutes.");
                            }
                            else
                            {
                                currentDate[commandElements[0]] = (toMinutes, currentDate[commandElements[0]].realMinutes);
                                Console.WriteLine($"Time in {commandElements[0]} set to: {toMinutes} minutes.");
                            }
                            break;
                        }
                        Console.WriteLine($"Invalid input! Type \"{commandElements[0]}\" to see all commands.");
                        break;
                    default:
                        int minutesChanged = 0;
                        if (commandElements.Length > 2 && int.TryParse(commandElements[2], out minutesChanged))
                        {
                            if (isRealTime)
                            {
                                currentDate[commandElements[0]] = (currentDate[commandElements[0]].minutes, currentDate[commandElements[0]].realMinutes + minutesChanged);
                                Console.WriteLine($"Real time in {commandElements[0]} changed by {(minutesChanged > 0 ? "+" : "")}{minutesChanged} minutes.");
                            }
                            else
                            {
                                currentDate[commandElements[0]] = (currentDate[commandElements[0]].minutes + minutesChanged, currentDate[commandElements[0]].realMinutes);
                                Console.WriteLine($"Time in {commandElements[0]} changed by {(minutesChanged > 0 ? "+" : "")}{minutesChanged} minutes.");
                            }
                            break;
                        }
                        Console.WriteLine($"Invalid input! Type \"{commandElements[0]}\" to see all subject commands.");
                        break;
                }
            }
            else
            {
                if (isRealTime)
                {
                    Console.WriteLine($"Real time spent on {commandElements[0]}: {ConvertMinutesToHoursString(currentDate[commandElements[0]].realMinutes)}.");
                }
                else
                {
                    Console.WriteLine($"Time spent on {commandElements[0]}: {ConvertMinutesToHoursString(currentDate[commandElements[0]].minutes)}.");
                }
            }
        }

        static void OnStartup()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;

            dataPath = Path.GetDirectoryName(Path.GetFullPath(Assembly.GetEntryAssembly().Location)) + @"\l_track.txt";

            if (!File.Exists(dataPath))
            {
                ProgramSetup();
            }

            LoadMetaData();

            SetCurrentDayOnStartup();

            WelcomeUser();
        }

        static int GetGradeInSetup()
        {
            Console.WriteLine("Which grade are you in?");
            string input = Console.ReadLine();
            ClearConsole(false);
            int grade = 0;
            while (!int.TryParse(input, out grade) || grade < 1 || grade > 12)
            {
                Console.WriteLine("Please enter a valid grade between 1 and 12 inclusive.");
                input = Console.ReadLine();
                ClearConsole(false);
            }
            return grade;
        }

        static (bool isInputValid, DateTime date) GetDateCheckInSetup()
        {
            Console.WriteLine($"When does/did the school year start?");
            Console.WriteLine($"Make sure the date format is \"dd/mm/yyyy\" - for example {DateTime.Now.Day}/{DateTime.Now.Month}/{DateTime.Now.Year}");
            string dateInput = Console.ReadLine();
            var dateCheckResult = DateStringToDateTime(dateInput);
            ClearConsole(false);
            while (!dateCheckResult.isInputValid)
            {
                Console.WriteLine("Please enter a valid date.");
                Console.WriteLine($"Make sure the date format is \"dd/mm/yyyy\" - for example \"{DateTime.Now.Day}/{DateTime.Now.Month}/{DateTime.Now.Year}\"");
                dateInput = Console.ReadLine();
                dateCheckResult = DateStringToDateTime(dateInput);
                ClearConsole(false);
            }
            return dateCheckResult;
        }

        public static string[] GetSubjectsInSetup()
        {
            Console.WriteLine($"What subjects do you have at school?");
            Console.WriteLine($"For example \"BEL,AE,NE,MAT\"");
            string[] subjects = Console.ReadLine().Split(',').Select(n => n.Trim()).ToArray();
            ClearConsole(false);
            return subjects;
        }

        public static bool IsUserInputCorrectInSetup(int grade, (bool isValid, DateTime date) dateCheckResult, string[] subjects)
        {

            Console.WriteLine($"You are in grade {grade}, which starts/has started on {dateCheckResult.date.Day}/{dateCheckResult.date.Month}/{dateCheckResult.date.Year}");
            Console.WriteLine($"The subjects you have at school are: {string.Join(", ", subjects)}");
            Console.WriteLine($"Is the information correct? Type \"yes\" if it is or \"no\" to restart the setup.");

            bool isCorrect = false;
            string queryAnswer = Console.ReadLine();
            ClearConsole(false);
            while (queryAnswer != "yes" && queryAnswer != "no")
            {
                Console.WriteLine($"You are in grade {grade}, which starts/has started on {dateCheckResult.date.Day}/{dateCheckResult.date.Month}/{dateCheckResult.date.Year}");
                Console.WriteLine($"The subjects you have at school are: {string.Join(", ", subjects)}");
                Console.WriteLine($"Is the information correct? Type \"yes\" if it is or \"no\" to restart the setup.");
                queryAnswer = Console.ReadLine();
                ClearConsole(false);
            }
            if (queryAnswer == "yes") { isCorrect = true; }
            return isCorrect;
        }

        static void ProgramSetup()
        {
            ClearConsole(false);

            int grade = GetGradeInSetup();

            var dateCheckResult = GetDateCheckInSetup();

            string[] subjects = GetSubjectsInSetup();

            bool isUserInputCorrect = IsUserInputCorrectInSetup(grade, dateCheckResult, subjects);

            if (isUserInputCorrect)
            {
                ProgramDataForYear newProgramDataForYear = new ProgramDataForYear(grade, subjects, dateCheckResult.date);
                ProgramDataMeta newProgramDataMeta = new ProgramDataMeta(grade, newProgramDataForYear);
                programDataForYear = newProgramDataForYear;
                metaData = newProgramDataMeta;
                ClearConsole(false);
                SaveMetaData();
                return;
            }
            else
            {
                ClearConsole(false);
                ProgramSetup();
            }
        }

        static int GetDaysFromStartToNow()
        {
            return (int)(DateTime.Now - programDataForYear.startDate).TotalDays;
        }

        static void LoadMetaData()
        {
            using (StreamReader sr = new StreamReader(dataPath))
            {
                metaData = JsonSerializer.Deserialize<ProgramDataMeta>(sr.ReadToEnd());
            }

            programDataForYear = metaData.programDataForAllYears.FirstOrDefault(n => n.grade == metaData.currentGrade);

            if (programDataForYear == null)
            {
                if (metaData.programDataForAllYears.Count > 0)
                {
                    int latestGrade = metaData.programDataForAllYears[metaData.programDataForAllYears.Count - 1].grade;
                    Console.WriteLine("Current grade set to " + latestGrade);
                    metaData.currentGrade = latestGrade;
                }
                else
                {
                    throw new NullReferenceException("metaData.programDataForAllYears is empty");
                }
            }
        }

        static void SaveMetaData()
        {
            using (StreamWriter sr = new StreamWriter(dataPath))
            {
                sr.Write(JsonSerializer.Serialize<ProgramDataMeta>(metaData));
            }
        }

        static void SetCurrentDayOnStartup()
        {
            int currentDayInDaysSinceStart = GetDaysFromStartToNow();

            SetCurrentDay(currentDayInDaysSinceStart);
        }

        static void SetCurrentDay(int daysSinceSchoolStarted)
        {
            if (daysSinceSchoolStarted < 0)
            {
                throw new ArgumentOutOfRangeException($"invalid day {daysSinceSchoolStarted}");
            }
            if (daysSinceSchoolStarted >= programDataForYear.leistungDates.Count)
            {
                CreateEmptyDaysTo(daysSinceSchoolStarted + 1);
            }
            currentDate = programDataForYear.leistungDates[daysSinceSchoolStarted];
        }

        static void CreateEmptyDaysTo(int toDaySinceSchoolStarted)
        {
            for (int i = programDataForYear.leistungDates.Count; i < toDaySinceSchoolStarted; i++)
            {
                programDataForYear.leistungDates.Add(new LeistungDate(i));
            }
        }

        public static void WriteColoredLine(string input, ConsoleColor cc)
        {
            ConsoleColor startColor = Console.ForegroundColor;
            Console.ForegroundColor = cc;
            Console.WriteLine(input);
            Console.ForegroundColor = startColor;
        }

        public static void WriteColored(string input, ConsoleColor cc)
        {
            ConsoleColor startColor = Console.ForegroundColor;
            Console.ForegroundColor = cc;
            Console.Write(input);
            Console.ForegroundColor = startColor;
        }

        public static string ConvertMinutesToHoursString(int minutesInput)
        {
            int hoursOutput = minutesInput / 60;
            int minutesOutput = minutesInput - hoursOutput * 60;

            return $"{(hoursOutput > 0 ? $"{hoursOutput}h" : "")} {minutesOutput}m".Trim();
        }

        public class ProgramDataForYear
        {
            public int grade { get; set; }
            public List<string> subjects { get; set; }
            public DateTime startDate { get; set; }
            public List<LeistungDate> leistungDates { get; set; }

            public ProgramDataForYear(int grade, IEnumerable<string> subjects, DateTime startDate)
            {
                this.grade = grade;
                this.subjects = subjects.ToList();
                this.startDate = startDate;
                this.leistungDates = new List<LeistungDate>();
            }

            public ProgramDataForYear(int grade, IEnumerable<string> subjects, DateTime startDate, IEnumerable<LeistungDate> leistungDates)
            {
                this.grade = grade;
                this.subjects = subjects.ToList();
                this.startDate = startDate;
                this.leistungDates = leistungDates.ToList();
            }

            public ProgramDataForYear() { }
        }

        public class ProgramDataMeta
        {
            public int currentGrade { get; set; }
            public List<ProgramDataForYear> programDataForAllYears { get; set; }

            public ProgramDataMeta(int currentGrade, params ProgramDataForYear[] programDataForAllYears)
            {
                this.currentGrade = currentGrade;
                this.programDataForAllYears = programDataForAllYears.ToList();
            }

            public ProgramDataMeta() { }
        }

        public class LeistungDate
        {
            public int daysSinceFirstSchoolDay { get; set; }
            public List<int> minutesPerSubject { get; set; }
            public List<int> realMinutesPerSubject { get; set; }
            private bool expanded = false;

            public (int minutes, int realMinutes) this[string subject]
            {
                get
                {
                    if (!expanded)
                    {
                        this.ExpandAllMinutesPerSubject();
                        expanded = true;
                    }
                    int subjectIndex = GetSubjectIndex(subject);
                    return (minutesPerSubject[subjectIndex], realMinutesPerSubject[subjectIndex]);
                }
                set
                {
                    if (!expanded)
                    {
                        this.ExpandAllMinutesPerSubject();
                        expanded = true;
                    }
                    int subjectIndex = GetSubjectIndex(subject);
                    this.minutesPerSubject[subjectIndex] = value.minutes;
                    this.realMinutesPerSubject[subjectIndex] = value.realMinutes;
                }
            }

            public LeistungDate() { }

            public LeistungDate(int daysSinceFirstSchoolDay)
            {
                realMinutesPerSubject = new List<int>();
                minutesPerSubject = new List<int>();
                this.daysSinceFirstSchoolDay = daysSinceFirstSchoolDay;
            }

            private int GetSubjectIndex(string subject)
            {
                int subjectIndex = programDataForYear.subjects.IndexOf(subject);
                if (subjectIndex == -1)
                {
                    throw new ArgumentOutOfRangeException("Invalid subject name.");
                }
                return subjectIndex;
            }

            private void ExpandAllMinutesPerSubject()
            {
                for (int i = minutesPerSubject.Count - 1; i < programDataForYear.subjects.Count; i++)
                {
                    minutesPerSubject.Add(0);
                }
                for (int i = realMinutesPerSubject.Count - 1; i < programDataForYear.subjects.Count; i++)
                {
                    realMinutesPerSubject.Add(0);
                }
            }
        }
    }
}