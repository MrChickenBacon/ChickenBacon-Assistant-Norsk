using System;
using System.Media;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;

namespace ChickenBaconAssistantBjarne
{
    class Program
    {
        [DllImport("User32", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(int uiAction, int uiParam,
            string pvParam, uint fWinIni);

        public static string Path = $"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}";
        private static readonly SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        public static SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        static Random random = new Random();

        static string[] Bjarnemat = { "Kebab", "Sandefjord Pizza", "McDonalds", "Thaimat", "Kinesisk", "Hva med en slik en, pizza grandiosa", "pølser i frysern", "Burger King" };

        static int Numbers()
        {
            int number = random.Next(1, 500);
            return number;
        }

        static void Main(string[] args)
        {
            synthesizer.SelectVoice("Microsoft Jon");
            SpeechRecoEngine();
            synthesizer.SpeakAsync("Hei, det er meg. Bjarne! Hva kan jeg gjøre for deg idag?");
            Thread.Sleep(1000);
            Console.WriteLine("Venter på stemme kommando.");
            Console.ReadKey();
        }

        private static void SpeechRecoEngine()
        {
            var gBuilder = CommandsGrammarBuilder();
            Grammar grammar = new Grammar(gBuilder);
            recEngine.LoadGrammarAsync(grammar);
            try
            {
                recEngine.SetInputToDefaultAudioDevice();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine();
                Console.WriteLine("Please plug in a microphone.");
                Console.ReadLine();
                throw;
            }
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        private static GrammarBuilder CommandsGrammarBuilder()
        {
            Choices commands = new Choices();
            commands.Add(new string[]
            {
                "Bjarne lag en test lyd", "kan du gi meg værmeldingen for idag",
                "hvem lagde deg", "hei Bjarne",
                "gi meg et tall", "ripsbusker og andre buskevekster", "Takk Bjarne",
                "klarer du o snakke engelsk", "skift bakgrunnsbilde for meg", "bjarne, si at ole skal stå opp nå", "gi meg et bjarnemat forslag"
            });
            //"Fortell meg litt om Get Academy",
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            return gBuilder;
        }

        private static void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "Bjarne lag en test lyd":
                    SoundPlayer player = new SoundPlayer($@"C:\Windows\media\tada.wav");
                    player.Play();
                    Thread.Sleep(2000);
                    synthesizer.SpeakAsync("Oj. Hørte du det? Det var en slags. TA.DA?");
                    Console.WriteLine("TADA");
                    break;
                case "hei Bjarne":
                    synthesizer.SpeakAsync("Hei på deg.");
                    Console.WriteLine("Hei");
                    break;
                case "kan du gi meg værmeldingen for idag":
                    synthesizer.SpeakAsync("Vet du hva? Du kan gå bort til et vindu og se ut. Jeg er ikke slaven din. Eller. Det er jeg.");
                    break;
                case "hvem lagde deg":
                    synthesizer.SpeakAsync("Det var vår store herre. Sjiken Baikken.");
                    break;
                case "gi meg et tall":
                    synthesizer.SpeakAsync("Du får " + Numbers());
                    break;
                case "ripsbusker og andre buskevekster":
                    synthesizer.SpeakAsync("Oj, det var en slik en tunge twisterino. Kult.");
                    break;
                case "klarer du o snakke engelsk":
                    synthesizer.SpeakAsync("Ehh. Jeg kan prøve. How are you today mister? Very much fine indeed? You takes wrong about the weather today. It is many cold degrees, and e is fresing my batteries off.");
                    break;
                case "skift bakgrunnsbilde for meg":
                    SystemParametersInfo(0x0014, 0, $@"{Path}\desktop\wallpapers\{Numbers()}.jpg", 0x0001);
                    synthesizer.SpeakAsync("Sånn. Jeg syns denne er fin.");
                    break;
                case "Takk Bjarne":
                    synthesizer.SpeakAsync("Det er bare hyggelig å hjelpe!");
                    Console.WriteLine(":)");
                    break;
                case "Fortell meg litt om Get Academy":
                    synthesizer.SpeakAsync("GET Academy AS leverer utdanning i IT-utvikling og nøkkelkompetanser. G'en i GET er Geir Sollid. E'en Eskil Domben. T'en Terje Kolderup. Skolen holder til i andre etasje på Torget 8 i Larvik.");
                    Console.WriteLine("http://www.getacademy.no/");
                    break;
                case "bjarne, si at ole skal stå opp nå":
                    synthesizer.SpeakAsync("Ok. Ole. Gunnar. Hellenes. Du skal stå opp nå. Hører du? Det er nudler med kylling på kjøkkenet. Du? Kanskje jeg også kan smake på det?");
                    break;
                case "gi meg et bjarnemat forslag":
                    synthesizer.SpeakAsync("Ja da skal vi se. Kanskje " + Bjarnemat[random.Next(0, Bjarnemat.Length)] + ", vil friste?");
                    break;
            }
        }
    }
}
