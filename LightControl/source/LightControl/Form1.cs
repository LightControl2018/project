using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.CognitiveServices.Speech;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Threading;
using System.Xml.Linq;
using System.Media;

namespace LightControl
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            button1.Text = "开始";
            pictureBox1.Load("lightoff.png");
            pictureBox2.Load("airconditionoff.png");
            pictureBox3.Load("lightoff.png");
            pictureBox4.Load("airconditionoff.png");
        }

        private static void PlayAudio(object sender, GenericEventArgs<Stream> args)
        {
            Console.WriteLine(args.EventData);

            SoundPlayer player = new SoundPlayer(args.EventData);
            player.PlaySync();
            args.EventData.Dispose();
        }
        private static void ErrorHandler(object sender, GenericEventArgs<Exception> e)
        {
            Console.WriteLine("Unable to complete the TTS request: [{0}]", e.ToString());
        }
        private void output(string txt)
        {
            Log("Starting Authtentication");
            string accessToken;
            Authentication auth = new Authentication("https://westus.api.cognitive.microsoft.com/sts/v1.0/issueToken", "688e38c1cb694c599c73815224dc6fbc");
            try
            {
                accessToken = auth.GetAccessToken();
                Log("Token:\n" + accessToken);
            }
            catch (Exception ex)
            {
                Log("Failed authentication.");
                Log(ex.ToString());
                Log(ex.Message);
                return;
            }
            Log("Starting TTSSample request code execution.");

            string requestUri = "https://westus.tts.speech.microsoft.com/cognitiveservices/v1";
            var cortana = new Synthesize();

            cortana.OnAudioAvailable += PlayAudio;
            cortana.OnError += ErrorHandler;

            cortana.Speak(CancellationToken.None, new Synthesize.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                // Text to be spoken.
                Text = txt,
                VoiceType = Gender.Female,
                // Refer to the documentation for complete list of supported locales.
                Locale = "zh-CN",
                // You can also customize the output voice. Refer to the documentation to view the different
                // voices that the TTS service can output.
                // VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, Jessa24KRUS)",
                VoiceName = "Microsoft Server Speech Text to Speech Voice (zh-CN, Yaoyao, Apollo)",
                // VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, ZiraRUS)",

                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }).Wait();

        }




        // 语音识别器
        SpeechRecognizer recognizer;
        bool isRecording = false;
        Dictionary<string, string> entities = new Dictionary<string, string>();
        Dictionary<string, int> ifexist = new Dictionary<string, int>();
        bool flag = false;//判断是否需要二次识别
        private void Form1_Load(object sender, EventArgs e)
        {
            entities.Add("location", "null");
            entities.Add("device", "null");

            try
            {
                // 第一步
                // 初始化语音服务SDK并启动识别器，进行语音转文本
                // 密钥和区域可在 https://azure.microsoft.com/zh-cn/try/cognitive-services/my-apis/?api=speech-services 中找到
                // 密钥示例: 5ee7ba6869f44321a40751967accf7a9
                // 区域示例: westus
                SpeechFactory speechFactory = SpeechFactory.FromSubscription("aee80335bb2049d593c906f5aa208b50", "westus");

                // 识别中文
                recognizer = speechFactory.CreateSpeechRecognizer("zh-CN");

                // 识别过程中的中间结果
                recognizer.IntermediateResultReceived += Recognizer_IntermediateResultReceived;
                // 识别的最终结果
                recognizer.FinalResultReceived += Recognizer_FinalResultReceived;
                // 出错时的处理
                recognizer.RecognitionErrorRaised += Recognizer_RecognitionErrorRaised;
            }
            catch (Exception ex)
            {
                if (ex is System.TypeInitializationException)
                {
                    Log("语音SDK不支持Any CPU, 请更改为x64");
                }
                else
                {
                    Log("初始化出错，请确认麦克风工作正常");
                    Log("已降级到文本语言理解模式");

                    TextBox inputBox = new TextBox();
                    inputBox.Text = "";
                    inputBox.Size = new Size(300, 26);
                    inputBox.Location = new Point(10, 10);
                    inputBox.KeyDown += inputBox_KeyDown;
                    Controls.Add(inputBox);

                    button1.Visible = false;
                }
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;

            isRecording = !isRecording;
            if (isRecording)
            {
                // 启动识别器
                await recognizer.StartContinuousRecognitionAsync();
                button1.Text = "停止";
            }
            else
            {
                // 停止识别器
                await recognizer.StopContinuousRecognitionAsync();
                button1.Text = "开始";
            }

            button1.Enabled = true;
        }

        // 识别过程中的中间结果
        private void Recognizer_IntermediateResultReceived(object sender, SpeechRecognitionResultEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Result.Text))
            {
                Log("中间结果: " + e.Result.Text);
            }
        }

        // 识别的最终结果
        private void Recognizer_FinalResultReceived(object sender, SpeechRecognitionResultEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Result.Text))
            {
                Log("最终结果: " + e.Result.Text);
                ProcessSttResult(e.Result.Text);
            }
        }

        // 出错时的处理
        private void Recognizer_RecognitionErrorRaised(object sender, RecognitionErrorEventArgs e)
        {
            Log("错误: " + e.FailureReason);
        }

        private async void ProcessSttResult(string text)
        {
            // 第二步
            // 调用语言理解服务取得用户意图

            string intent = await GetLuisResult(text);
            if (flag)
            {
                intent = await GetLuisResult(intent);
                flag = false;
            }
            // 第三步
            // 按照意图控制灯

            if (!string.IsNullOrEmpty(intent))
            {
                if (intent.Equals("turn kitchen light on", StringComparison.OrdinalIgnoreCase))
                {

                    OpenKitchenLight();
                    //输出
                    output("厨房灯已开呢");

                }
                else if (intent.Equals("turn the kitchen light off", StringComparison.OrdinalIgnoreCase))
                {

                    CloseKitchenLight();
                    //输出
                    output("厨房灯已关喽");
                }
                else if (intent.Equals("turn kitchen air conditioner on", StringComparison.OrdinalIgnoreCase))
                {
                    OpenKitchenAircondition();
                    //输出
                    output("厨房空调开啦");
                }
                else if (intent.Equals("turn air conditioner in kitchen off", StringComparison.OrdinalIgnoreCase))
                {
                    CloseKitchenAircondition();
                    //输出
                    output("关上厨房空调");

                }
                else if (intent.Equals("turn toilet light on", StringComparison.OrdinalIgnoreCase))
                {
                    OpenWcLight();
                    //输出
                    output("厕所灯已开开");

                }
                else if (intent.Equals("turn toilet light off", StringComparison.OrdinalIgnoreCase))
                {
                    CloseWcLight();
                    //输出
                    output("厕所灯已关关");

                }
                else if (intent.Equals("turn toilet air conditioner on", StringComparison.OrdinalIgnoreCase))
                {
                    OpenWcAircondition();
                    //输出
                    output("厕所空调已开开呢啦");
                }
                else if (intent.Equals("turn toilet air conditioner off", StringComparison.OrdinalIgnoreCase))
                {
                    CloseWcAircondition();
                    //输出
                    output("厕所空调shut down");
                }
            }

        }



        // 第二步
        // 调用语言理解服务取得用户意图

        private async Task<string> GetLuisResult(string text)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                // LUIS 终结点地址, 示例: https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/102f6255-0c32-4f36-9c79-fe12fea4d6c4?subscription-key=9004421650254a74876cf3c888b1d11f&verbose=true&timezoneOffset=0&q=
                // 可在 https://www.luis.ai 中进入app右上角publish中找到
                string luisEndpoint = "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/08374698-5ce8-4316-8058-bae2593ac761?subscription-key=7da65785f261488e9f8820e8db4adbff&verbose=true&timezoneOffset=0&q= ";
                string luisJson = await httpClient.GetStringAsync(luisEndpoint + text);

                try
                {
                    dynamic result = JsonConvert.DeserializeObject<dynamic>(luisJson);
                    string intent = (string)result.topScoringIntent.intent;
                    double score = 0;
                    int size = result.entities.Count;
                    if (size == 0)
                    {
                        Log("illegal!\r\n");
                        return "illegal!";
                    }
                    else if (size == 2)
                    {
                        score = (double)result.topScoringIntent.score;
                        Log("意图: " + intent + "\r\n得分: " + score + "\r\n");
                        entities["location"] = result.entities[1].entity;
                        entities["device"] = result.entities[0].entity;

                        return intent;
                    }
                    else
                    {
                        flag = true;
                        if (result.entities[0].type == "location")
                        {
                            entities["location"] = result.entities[0].entity;

                            text = text + entities["device"];
                            return text;
                        }
                        else
                        {
                            entities["device"] = result.entities[0].entity;
                            text = text + entities["location"];
                            return text;
                        }
                    }




                }
                catch (Exception ex)
                {
                    Log(ex.Message);
                    return null;
                }
            }
        }


        #region 界面操作

        private void Log(string message)
        {
            MakesureRunInUI(() =>
            {
                textBox1.AppendText(message + "\r\n");
            });
        }

        private void OpenWcLight()
        {
            MakesureRunInUI(() =>
            {
                pictureBox1.Load("lighton.png");
            });
        }
        private void OpenWcAircondition()
        {
            MakesureRunInUI(() =>
            {
                pictureBox2.Load("airconditionon.png");
            });
        }
        private void OpenKitchenLight()
        {
            MakesureRunInUI(() =>
            {
                pictureBox3.Load("lighton.png");
            });
        }
        private void OpenKitchenAircondition()
        {
            MakesureRunInUI(() =>
            {
                pictureBox4.Load("airconditionon.png");
            });
        }
        private void CloseWcLight()
        {
            MakesureRunInUI(() =>
            {
                pictureBox1.Load("lightoff.png");
            });
        }
        private void CloseWcAircondition()
        {
            MakesureRunInUI(() =>
            {
                pictureBox2.Load("airconditionoff.png");
            });
        }
        private void CloseKitchenLight()
        {
            MakesureRunInUI(() =>
            {
                pictureBox3.Load("lightoff.png");
            });
        }
        private void CloseKitchenAircondition()
        {
            MakesureRunInUI(() =>
            {
                pictureBox4.Load("airconditionoff.png");
            });
        }
        private void MakesureRunInUI(Action action)
        {
            if (InvokeRequired)
            {
                MethodInvoker method = new MethodInvoker(action);
                Invoke(action, null);
            }
            else
            {
                action();
            }
        }

        #endregion

        private void inputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && sender is TextBox)
            {
                TextBox textBox = sender as TextBox;
                e.Handled = true;
                Log(textBox.Text);
                ProcessSttResult(textBox.Text);

                textBox.Text = string.Empty;
            }
        }

    }
    public class Authentication
    {
        private string AccessUri;
        private string apiKey;
        private string accessToken;
        private System.Threading.Timer accessTokenRenewer;

        private const int RefreshTokenDuration = 9;

        public Authentication(string issueTokenUri, string apiKey)
        {
            this.AccessUri = issueTokenUri;
            this.apiKey = apiKey;

            this.accessToken = HttpPost(issueTokenUri, this.apiKey);

            accessTokenRenewer = new System.Threading.Timer(new TimerCallback(OnTokenExpiredCallback),
                                           this,
                                           TimeSpan.FromMinutes(RefreshTokenDuration),
                                           TimeSpan.FromMilliseconds(-1));
        }

        public string GetAccessToken()
        {
            return this.accessToken;
        }

        private void RenewAccessToken()
        {
            string newAccessToken = HttpPost(AccessUri, this.apiKey);
            this.accessToken = newAccessToken;
            Console.WriteLine(string.Format("Renewed token for user: {0} is: {1}",
                              this.apiKey,
                              this.accessToken));
        }

        private void OnTokenExpiredCallback(object stateInfo)
        {
            try
            {
                RenewAccessToken();
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Failed renewing access token. Details: {0}", ex.Message));
            }
            finally
            {
                try
                {
                    accessTokenRenewer.Change(TimeSpan.FromMinutes(RefreshTokenDuration), TimeSpan.FromMilliseconds(-1));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("Failed to reschedule the timer to renew access token. Details: {0}", ex.Message));
                }
            }
        }

        private string HttpPost(string accessUri, string apiKey)
        {
            WebRequest webRequest = WebRequest.Create(accessUri);
            webRequest.Method = "POST";
            webRequest.ContentLength = 0;
            webRequest.Headers["Ocp-Apim-Subscription-Key"] = apiKey;

            using (WebResponse webResponse = webRequest.GetResponse())
            {
                using (Stream stream = webResponse.GetResponseStream())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        byte[] waveBytes = null;
                        int count = 0;
                        do
                        {
                            byte[] buf = new byte[1024];
                            count = stream.Read(buf, 0, 1024);
                            ms.Write(buf, 0, count);
                        } while (stream.CanRead && count > 0);

                        waveBytes = ms.ToArray();

                        return Encoding.UTF8.GetString(waveBytes);
                    }
                }
            }
        }
    }
    public class GenericEventArgs<T> : EventArgs
    {
        public GenericEventArgs(T eventData)
        {
            this.EventData = eventData;
        }

        public T EventData { get; private set; }
    }

    public enum Gender
    {
        Female,
        Male
    }

    public enum AudioOutputFormat
    {
        Raw8Khz8BitMonoMULaw,
        Raw16Khz16BitMonoPcm,
        Riff8Khz8BitMonoMULaw,
        Riff16Khz16BitMonoPcm,
        Ssml16Khz16BitMonoSilk,
        Raw16Khz16BitMonoTrueSilk,
        Ssml16Khz16BitMonoTts,
        Audio16Khz128KBitRateMonoMp3,
        Audio16Khz64KBitRateMonoMp3,
        Audio16Khz32KBitRateMonoMp3,
        Audio16Khz16KbpsMonoSiren,
        Riff16Khz16KbpsMonoSiren,
        Raw24Khz16BitMonoTrueSilk,
        Raw24Khz16BitMonoPcm,
        Riff24Khz16BitMonoPcm,
        Audio24Khz48KBitRateMonoMp3,
        Audio24Khz96KBitRateMonoMp3,
        Audio24Khz160KBitRateMonoMp3
    }
    public class Synthesize
    {
        private string GenerateSsml(string locale, string gender, string name, string text)
        {
            var ssmlDoc = new XDocument(
                              new XElement("speak",
                                  new XAttribute("version", "1.0"),
                                  new XAttribute(XNamespace.Xml + "lang", "en-US"),
                                  new XElement("voice",
                                      new XAttribute(XNamespace.Xml + "lang", locale),
                                      new XAttribute(XNamespace.Xml + "gender", gender),
                                      new XAttribute("name", name),
                                      text)));
            return ssmlDoc.ToString();
        }

        private HttpClient client;
        private HttpClientHandler handler;

        public Synthesize()
        {
            var cookieContainer = new CookieContainer();
            handler = new HttpClientHandler() { CookieContainer = new CookieContainer(), UseProxy = false };
            client = new HttpClient(handler);
        }

        ~Synthesize()
        {
            client.Dispose();
            handler.Dispose();
        }

        public event EventHandler<GenericEventArgs<Stream>> OnAudioAvailable;

        public event EventHandler<GenericEventArgs<Exception>> OnError;
        public Task Speak(CancellationToken cancellationToken, InputOptions inputOptions)
        {
            client.DefaultRequestHeaders.Clear();
            foreach (var header in inputOptions.Headers)
            {
                client.DefaultRequestHeaders.TryAddWithoutValidation(header.Key, header.Value);
            }

            var genderValue = "";
            switch (inputOptions.VoiceType)
            {
                case Gender.Male:
                    genderValue = "Male";
                    break;

                case Gender.Female:
                default:
                    genderValue = "Female";
                    break;
            }

            var request = new HttpRequestMessage(HttpMethod.Post, inputOptions.RequestUri)
            {
                Content = new StringContent(GenerateSsml(inputOptions.Locale, genderValue, inputOptions.VoiceName, inputOptions.Text))
            };

            var httpTask = client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            Console.WriteLine("Response status code: [{0}]", httpTask.Result.StatusCode);

            var saveTask = httpTask.ContinueWith(
                async (responseMessage, token) =>
                {
                    try
                    {
                        if (responseMessage.IsCompleted && responseMessage.Result != null && responseMessage.Result.IsSuccessStatusCode)
                        {
                            var httpStream = await responseMessage.Result.Content.ReadAsStreamAsync().ConfigureAwait(false);
                            this.AudioAvailable(new GenericEventArgs<Stream>(httpStream));
                        }
                        else
                        {
                            this.Error(new GenericEventArgs<Exception>(new Exception(String.Format("Service returned {0}", responseMessage.Result.StatusCode))));
                        }
                    }
                    catch (Exception e)
                    {
                        this.Error(new GenericEventArgs<Exception>(e.GetBaseException()));
                    }
                    finally
                    {
                        responseMessage.Dispose();
                        request.Dispose();
                    }
                },
                TaskContinuationOptions.AttachedToParent,
                cancellationToken);

            return saveTask;
        }
        private void AudioAvailable(GenericEventArgs<Stream> e)
        {
            EventHandler<GenericEventArgs<Stream>> handler = this.OnAudioAvailable;
            if (handler != null)
            {
                handler(this, e);
            }
        }
        private void Error(GenericEventArgs<Exception> e)
        {
            EventHandler<GenericEventArgs<Exception>> handler = this.OnError;
            if (handler != null)
            {
                handler(this, e);
            }
        }
        public class InputOptions
        {
            public InputOptions()
            {
                this.Locale = "zh-cn";
                this.VoiceName = "Microsoft Server Speech Text to Speech Voice (zh-CN, Yaoyao, Apollo)";
                this.OutputFormat = AudioOutputFormat.Riff16Khz16BitMonoPcm;
            }
            public Uri RequestUri { get; set; }
            public AudioOutputFormat OutputFormat { get; set; }
            public IEnumerable<KeyValuePair<string, string>> Headers
            {
                get
                {
                    List<KeyValuePair<string, string>> toReturn = new List<KeyValuePair<string, string>>();
                    toReturn.Add(new KeyValuePair<string, string>("Content-Type", "application/ssml+xml"));

                    string outputFormat;

                    switch (this.OutputFormat)
                    {
                        case AudioOutputFormat.Raw16Khz16BitMonoPcm:
                            outputFormat = "raw-16khz-16bit-mono-pcm";
                            break;

                        case AudioOutputFormat.Raw8Khz8BitMonoMULaw:
                            outputFormat = "raw-8khz-8bit-mono-mulaw";
                            break;

                        case AudioOutputFormat.Riff16Khz16BitMonoPcm:
                            outputFormat = "riff-16khz-16bit-mono-pcm";
                            break;

                        case AudioOutputFormat.Riff8Khz8BitMonoMULaw:
                            outputFormat = "riff-8khz-8bit-mono-mulaw";
                            break;

                        case AudioOutputFormat.Ssml16Khz16BitMonoSilk:
                            outputFormat = "ssml-16khz-16bit-mono-silk";
                            break;

                        case AudioOutputFormat.Raw16Khz16BitMonoTrueSilk:
                            outputFormat = "raw-16khz-16bit-mono-truesilk";
                            break;

                        case AudioOutputFormat.Ssml16Khz16BitMonoTts:
                            outputFormat = "ssml-16khz-16bit-mono-tts";
                            break;

                        case AudioOutputFormat.Audio16Khz128KBitRateMonoMp3:
                            outputFormat = "audio-16khz-128kbitrate-mono-mp3";
                            break;

                        case AudioOutputFormat.Audio16Khz64KBitRateMonoMp3:
                            outputFormat = "audio-16khz-64kbitrate-mono-mp3";
                            break;

                        case AudioOutputFormat.Audio16Khz32KBitRateMonoMp3:
                            outputFormat = "audio-16khz-32kbitrate-mono-mp3";
                            break;

                        case AudioOutputFormat.Audio16Khz16KbpsMonoSiren:
                            outputFormat = "audio-16khz-16kbps-mono-siren";
                            break;

                        case AudioOutputFormat.Riff16Khz16KbpsMonoSiren:
                            outputFormat = "riff-16khz-16kbps-mono-siren";
                            break;
                        case AudioOutputFormat.Raw24Khz16BitMonoPcm:
                            outputFormat = "raw-24khz-16bit-mono-pcm";
                            break;
                        case AudioOutputFormat.Riff24Khz16BitMonoPcm:
                            outputFormat = "riff-24khz-16bit-mono-pcm";
                            break;
                        case AudioOutputFormat.Audio24Khz48KBitRateMonoMp3:
                            outputFormat = "audio-24khz-48kbitrate-mono-mp3";
                            break;
                        case AudioOutputFormat.Audio24Khz96KBitRateMonoMp3:
                            outputFormat = "audio-24khz-96kbitrate-mono-mp3";
                            break;
                        case AudioOutputFormat.Audio24Khz160KBitRateMonoMp3:
                            outputFormat = "audio-24khz-160kbitrate-mono-mp3";
                            break;
                        default:
                            outputFormat = "riff-16khz-16bit-mono-pcm";
                            break;
                    }

                    toReturn.Add(new KeyValuePair<string, string>("X-Microsoft-OutputFormat", outputFormat));
                    toReturn.Add(new KeyValuePair<string, string>("Authorization", this.AuthorizationToken));
                    toReturn.Add(new KeyValuePair<string, string>("X-Search-AppId", "07D3234E49CE426DAA29772419F436CA"));
                    toReturn.Add(new KeyValuePair<string, string>("X-Search-ClientID", "1ECFAE91408841A480F00935DC390960"));
                    toReturn.Add(new KeyValuePair<string, string>("User-Agent", "TTSClient"));
                    return toReturn;
                }
                set
                {
                    Headers = value;
                }
            }
            public String Locale { get; set; }
            public Gender VoiceType { get; set; }
            public string VoiceName { get; set; }
            public string AuthorizationToken { get; set; }
            public string Text { get; set; }
        }
    }
}