using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using BepInEx;
using BepInEx.Configuration;
using BepInEx.Logging;
using HarmonyLib;
using UnityEngine;

namespace Empress.MoonbasePaul
{
    [BepInPlugin(PluginGuid, PluginName, PluginVersion)]
    public class MoonbasePaul : BaseUnityPlugin
    {
        public const string PluginGuid = "Empress.MoonbasePaul";
        public const string PluginName = "MoonbasePaul";
        public const string PluginVersion = "2.0.3";

        internal static MoonbasePaul Instance = null!;
        internal new static ManualLogSource Logger => Instance._logger;
        private ManualLogSource _logger => base.Logger;
        internal Harmony? Harmony { get; set; }
        internal static ConfigEntry<bool> Enabled = null!;
        internal static ConfigEntry<bool> LocalOnly = null!;
        internal static ConfigEntry<float> Wet = null!;
        internal static ConfigEntry<bool> ForceGameTTS = null!;
        internal static ConfigEntry<bool> UseDECtalk = null!;
        internal static ConfigEntry<string> DecTalkExePath = null!;        
        internal static ConfigEntry<string> DecTalkVoice = null!;          
        internal static ConfigEntry<int> DecTalkSampleRate = null!;        
        internal static ConfigEntry<string> DecTalkDictionaryPath = null!; 
        internal static ConfigEntry<bool> DecTalkWhisperOnCrouch = null!; 

        private Coroutine? _watcher;

        private static string PluginDir
        {
            get
            {
                var loc = Assembly.GetExecutingAssembly().Location;
                return string.IsNullOrEmpty(loc) ? Paths.PluginPath : (Path.GetDirectoryName(loc) ?? Paths.PluginPath);
            }
        }

        private void Awake()
        {
            Instance = this;
            this.gameObject.transform.parent = null;
            this.gameObject.hideFlags = HideFlags.HideAndDontSave;

            
            Enabled = Config.Bind("General", "Enabled", true, "Master toggle.");
            LocalOnly = Config.Bind("General", "LocalOnly", false, "Only affect your own TTS voice.");
            Wet = Config.Bind("Mix", "Wet", 1.0f, "Reserved for any post-mix (0..1). Currently no post effect applied.");
            ForceGameTTS = Config.Bind("General", "ForceGameTTS", false, "Bypass this mod and always use the game's built-in TTS.");

            
            UseDECtalk = Config.Bind("DECtalk", "UseDECtalk", true, "Use DECtalk (Perfect Paul) instead of the game's synth.");
            DecTalkExePath = Config.Bind("DECtalk", "ExePath", "say.exe", "Path or filename of DECtalk SAY.exe. Relative to the plugin DLL folder if not absolute.");
            DecTalkVoice = Config.Bind("DECtalk", "Voice", "Paul", "Voice name, e.g., Paul (Perfect Paul).");
            DecTalkSampleRate = Config.Bind("DECtalk", "SampleRate", 22050, "Target sample rate for clips (Hz). 11025/16000/22050/44100 typically fine.");
            DecTalkDictionaryPath = Config.Bind("DECtalk", "DictionaryPath", "dtalk_us.dic", "Path or filename of DECtalk main dictionary. Leave blank to omit -d.");
            DecTalkWhisperOnCrouch = Config.Bind("DECtalk", "WhisperOnCrouch", false, "If true, crouch uses Wendy ([:name Wendy]) as \"whisper\".");

            Harmony ??= new Harmony(PluginGuid);
            Harmony.PatchAll();

            _watcher = StartCoroutine(AttachLoop());
            Logger.LogInfo($"{Info.Metadata.GUID} v{Info.Metadata.Version} loaded. BaseDir={PluginDir}");
        }

        private void OnDestroy()
        {
            try { Harmony?.UnpatchSelf(); } catch { /* whatever */ }
            if (_watcher != null) StopCoroutine(_watcher);
        }

        private IEnumerator AttachLoop()
        {
            var seen = new HashSet<AudioSource>();
            var wait = new WaitForSeconds(0.5f);

            while (true)
            {
                if (Enabled.Value)
                {
                    var monos = GameObject.FindObjectsOfType<MonoBehaviour>();
                    foreach (var mb in monos)
                    {
                        if (!mb) continue;
                        var t = mb.GetType();
                        if (t.Name != "PlayerVoiceChat") continue;

                        var f = t.GetField("ttsAudioSource", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        if (f == null) continue;

                        if (!(f.GetValue(mb) is AudioSource audio)) continue;

                        if (LocalOnly.Value && !IsLocalPlayerVoice(mb))
                            continue;

                        if (seen.Add(audio))
                            Logger.LogInfo($"MoonbasePaul: hooked TTS AudioSource on {audio.gameObject.name}");
                    }
                }
                yield return wait;
            }
        }

        private static bool IsLocalPlayerVoice(MonoBehaviour playerVoiceChat)
        {
            var t = playerVoiceChat.GetType();
            var favatar = t.GetField("playerAvatar", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (favatar == null) return false;
            var avatar = favatar.GetValue(playerVoiceChat);
            if (avatar == null) return false;

            var isLocalProp = avatar.GetType().GetProperty("isLocal", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (isLocalProp != null && isLocalProp.PropertyType == typeof(bool))
                return (bool)(isLocalProp.GetValue(avatar) ?? false);

            var isLocalField = avatar.GetType().GetField("isLocal", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (isLocalField != null && isLocalField.FieldType == typeof(bool))
                return (bool)(isLocalField.GetValue(avatar) ?? false);

            return false;
        }

        internal static class DecTalk
        {
            internal struct Clip
            {
                public float[] Samples;
                public int Channels;
                public int SampleRate;
            }

            private static string ResolvePath(string? p)
            {
                if (string.IsNullOrWhiteSpace(p)) return string.Empty;
                return Path.IsPathRooted(p!) ? p! : Path.GetFullPath(Path.Combine(PluginDir, p!));
            }

            internal static bool TrySynthesize(string text, bool crouch, out Clip clip)
            {
                clip = default;

                try
                {
                    if (!UseDECtalk.Value || ForceGameTTS.Value)
                        return false;

                    var exe = ResolvePath(DecTalkExePath.Value.Trim());
                    if (string.IsNullOrEmpty(exe) || !File.Exists(exe))
                    {
                        Logger.LogWarning($"DECtalk SAY.exe not found at: {exe}");
                        return false;
                    }

                    var voice = (crouch && DecTalkWhisperOnCrouch.Value) ? "Wendy" : DecTalkVoice.Value;
                    var payload = $"[:name {voice}] {text}";

                    var exeDir = Path.GetDirectoryName(exe) ?? PluginDir;
                    var outDir = Path.Combine(exeDir, "out");
                    Directory.CreateDirectory(outDir);
                    var token = DateTime.UtcNow.Ticks.ToString("x");
                    var outWav = Path.Combine(outDir, $"dt_{token}.wav");

                    var dictArg = string.Empty;
                    var dictCfg = DecTalkDictionaryPath.Value?.Trim();
                    if (!string.IsNullOrEmpty(dictCfg))
                    {
                        var dictPath = ResolvePath(dictCfg);
                        if (File.Exists(dictPath))
                            dictArg = $"-d \"{dictPath}\" ";
                        else
                            Logger.LogWarning($"DECtalk dictionary not found at: {dictPath}. Proceeding without -d.");
                    }

                    var psi = new ProcessStartInfo
                    {
                        FileName = exe,
                        Arguments = $"-w \"{outWav}\" {dictArg}\"{payload}\"",
                        WorkingDirectory = exeDir,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    string? stdOut = null, stdErr = null;
                    using (var p = Process.Start(psi))
                    {
                        stdOut = p?.StandardOutput.ReadToEnd();
                        stdErr = p?.StandardError.ReadToEnd();
                        p?.WaitForExit(15000);
                        if (p != null && p.HasExited && p.ExitCode != 0)
                            Logger.LogWarning($"DECtalk exit code {p.ExitCode}. stderr: {stdErr}");
                    }

                    if (!File.Exists(outWav))
                    {
                        Logger.LogWarning($"DECtalk did not produce a WAV file: {outWav}. stderr: {stdErr}");
                        return false;
                    }

                    if (!Wav.TryLoad(outWav, out var samples, out var channels, out var sr))
                    {
                        Logger.LogWarning("Failed to parse DECtalk WAV.");
                        return false;
                    }

                    var targetSr = Mathf.Clamp(DecTalkSampleRate.Value, 8000, 48000);
                    if (sr != targetSr)
                    {
                        samples = Wav.ResampleLinear(samples, channels, sr, targetSr);
                        sr = targetSr;
                    }

                    clip = new Clip { Samples = samples, Channels = channels, SampleRate = sr };

                    try { File.Delete(outWav); } catch { }
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.LogError($"DECtalk synth failed: {ex}");
                    return false;
                }
            }

            private static class Wav
            {
                internal static bool TryLoad(string path, out float[] data, out int channels, out int sampleRate)
                {
                    data = Array.Empty<float>();
                    channels = 0;
                    sampleRate = 0;

                    using var fs = File.OpenRead(path);
                    using var br = new BinaryReader(fs);

                    if (new string(br.ReadChars(4)) != "RIFF") return false;
                    br.ReadInt32(); 
                    if (new string(br.ReadChars(4)) != "WAVE") return false;

                    short audioFormat = 1; 
                    short bitsPerSample = 16;
                    bool fmtSeen = false;

                    while (br.BaseStream.Position < br.BaseStream.Length)
                    {
                        var chunkId = new string(br.ReadChars(4));
                        var chunkSize = br.ReadInt32();
                        if (chunkId == "fmt ")
                        {
                            fmtSeen = true;
                            audioFormat = br.ReadInt16();
                            channels = br.ReadInt16();
                            sampleRate = br.ReadInt32();
                            br.ReadInt32(); 
                            br.ReadInt16(); 
                            bitsPerSample = br.ReadInt16();
                            var extra = chunkSize - 16;
                            if (extra > 0) br.ReadBytes(extra);
                        }
                        else if (chunkId == "data")
                        {
                            if (!fmtSeen) return false;
                            var bytes = br.ReadBytes(chunkSize);
                            if (audioFormat != 1) return false; 

                            if (bitsPerSample == 16)
                            {
                                int sampleCount = bytes.Length / 2;
                                var floats = new float[sampleCount];
                                for (int i = 0, j = 0; i < sampleCount; i++, j += 2)
                                {
                                    short s = (short)(bytes[j] | (bytes[j + 1] << 8));
                                    floats[i] = s / 32768f;
                                }
                                data = floats;
                                return true;
                            }
                            else if (bitsPerSample == 8)
                            {
                                int sampleCount = bytes.Length;
                                var floats = new float[sampleCount];
                                for (int i = 0; i < sampleCount; i++)
                                    floats[i] = (bytes[i] - 128) / 128f;
                                data = floats;
                                return true;
                            }
                            else return false;
                        }
                        else
                        {
                            br.ReadBytes(chunkSize);
                        }
                    }
                    return false;
                }

                internal static float[] ResampleLinear(float[] interleaved, int channels, int srcRate, int dstRate)
                {
                    if (srcRate == dstRate) return interleaved;
                    int frames = interleaved.Length / channels;
                    double ratio = (double)dstRate / srcRate;
                    int dstFrames = Mathf.Max(1, (int)Math.Round(frames * ratio));
                    var dst = new float[dstFrames * channels];

                    for (int c = 0; c < channels; c++)
                    {
                        for (int i = 0; i < dstFrames; i++)
                        {
                            double srcPos = i / ratio;
                            int i0 = (int)Math.Floor(srcPos);
                            int i1 = Math.Min(frames - 1, i0 + 1);
                            double t = srcPos - i0;
                            float s0 = interleaved[i0 * channels + c];
                            float s1 = interleaved[i1 * channels + c];
                            dst[i * channels + c] = (float)(s0 + (s1 - s0) * t);
                        }
                    }
                    return dst;
                }
            }
        }
    }

    [HarmonyPatch]
    internal static class Patch_TTSVoice_DecTalk
    {
        private static MethodBase TargetMethod()
        {
            var t = AccessTools.TypeByName("TTSVoice");
            return AccessTools.Method(t, "StartSpeakingWithHighlight", new Type[] { typeof(string), typeof(bool) });
        }

        private static bool Prefix(object __instance, string text, bool crouch)
        {
            try
            {
                if (!MoonbasePaul.Enabled.Value) return true;
                if (MoonbasePaul.ForceGameTTS.Value) return true;
                if (!MoonbasePaul.UseDECtalk.Value) return true;

                if (MoonbasePaul.LocalOnly.Value)
                {
                    var pavF = __instance.GetType().GetField("playerAvatar", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    var pav = pavF?.GetValue(__instance);
                    if (pav != null)
                    {
                        var isLocalProp = pav.GetType().GetProperty("isLocal", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        var isLocalField = pav.GetType().GetField("isLocal", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        bool isLocal = false;
                        if (isLocalProp?.PropertyType == typeof(bool)) isLocal = (bool)(isLocalProp.GetValue(pav) ?? false);
                        else if (isLocalField?.FieldType == typeof(bool)) isLocal = (bool)(isLocalField.GetValue(pav) ?? false);
                        if (!isLocal) return true;
                    }
                }

                if (!MoonbasePaul.DecTalk.TrySynthesize(text, crouch, out var clip))
                    return true; 

                var go = (__instance as Component)?.gameObject;
                if (!go) return true;
                var src = go.GetComponent<AudioSource>() ?? go.AddComponent<AudioSource>();

                var unityClip = AudioClip.Create("DECtalk", clip.Samples.Length / clip.Channels, clip.Channels, clip.SampleRate, false);
                unityClip.SetData(clip.Samples, 0);
                src.clip = unityClip;
                src.loop = false;
                src.Play();

                var voiceTextM = __instance.GetType().GetMethod("VoiceText", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                voiceTextM?.Invoke(__instance, new object[] { text, 0f });

                return false;
            }
            catch (Exception ex)
            {
                MoonbasePaul.Logger.LogError($"DECtalk patch failed: {ex}");
                return true;
            }
        }
    }
}