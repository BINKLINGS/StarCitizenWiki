# C:\Users\Lenovo\AppData\Local\Programs\Python\Python312\Lib\site-packages\pythonwin\Pythonwin.exe
# pyinstaller --onefile --hidden-import "win32com.gen_py.C866CA3A-32F7-11D2-9602-00C04F8EE628x0x5x4" voice_commander.py

import pythoncom
import win32com.client
from win32com.client import constants
import json
import pydirectinput
import time
import os
import threading
import sys # 用于获取 PyInstaller 的运行时路径



# 检查是否在打包环境中 (可选但推荐)
is_frozen = getattr(sys, 'frozen', False)

# SAPI 的类型库信息
# CLSID: {C866CA3A-32F7-11D2-9602-00C04F8EE628}, Version: 5.4
SAPI_CLSID = '{C866CA3A-32F7-11D2-9602-00C04F8EE628}'
SAPI_VERSION_MAJOR = 5
SAPI_VERSION_MINOR = 4
SAPI_LCID = 0  # Language Neutral

# 强制生成或确保 SAPI 的 COM 包装模块存在
# 这行代码是关键！它等同于在命令行运行 makepy.py
try:
    win32com.client.gencache.EnsureModule(SAPI_CLSID, SAPI_LCID, SAPI_VERSION_MAJOR, SAPI_VERSION_MINOR)
except Exception as e:
    print(f"无法生成 SAPI 的 COM 包装: {e}")
    # 在这里可以决定是退出程序还是继续尝试
    # 如果没有这个包装，getevents 几乎肯定会失败
    if is_frozen:
        # 在打包后的程序中，如果生成失败，可能需要提示用户并退出
        # 因为可能没有写入临时目录的权限
        input("错误：无法初始化语音识别组件。按回车键退出。")
        sys.exit(1)



# (list_sapi_resources 函数保持不变，为节省篇幅此处省略，但请确保你的脚本中包含它)
def list_sapi_resources():
    """打印出系统中所有可用的SAPI语音（输出）和识别器（输入）。"""
    print("--- 可用的语音 (输出) ---")
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        voices = speaker.GetVoices()
        print(f"找到 {voices.Count} 个语音:")
        for i, voice in enumerate(voices):
            print(f"  {i+1}. 名称: {voice.GetDescription()}")
    except Exception as e:
        print(f"无法获取语音列表: {e}")

    print("\n--- 可用的识别器 (输入) ---")
    try:
        cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
        cat.SetId(constants.SpeechCategoryRecognizers)
        recognizers = cat.EnumerateTokens()
        print(f"找到 {recognizers.Count} 个识别器:")
        for i, reco in enumerate(recognizers):
            print(f"  {i+1}. 名称: {reco.GetDescription()}")
    except Exception as e:
        print(f"无法获取识别器列表: {e}")
    print("-" * 30)

class SpeechRecognition:
    def __init__(self, config_data, voice_name=None, recognizer_name=None, speech_rate=0, initial_greeting=""):
        self.config_data = config_data
        wordsToAdd = [command['trigger'] for command in self.config_data.get('commands', [])]
        if not wordsToAdd:
            raise ValueError("错误：配置文件中未找到任何有效的 'commands' 或 'trigger'。")
        
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        if voice_name:
            self._set_voice(voice_name)
        
        self.speaker.Rate = speech_rate 
        print(f"TTS 语速已设置为: {self.speaker.Rate}")

        if initial_greeting:
            self.say(initial_greeting)

        self.listener = win32com.client.Dispatch("SAPI.SpInprocRecognizer")
        if recognizer_name:
            self._set_recognizer(recognizer_name)

        try:
            if self.listener.GetAudioInputs().Count == 0:
                raise RuntimeError("错误：未找到任何麦克风/音频输入设备。")
            self.listener.AudioInput = self.listener.GetAudioInputs().Item(0)
        except Exception as e:
            print(f"设置音频输入时出错: {e}")
            self.say("初始化音频输入失败")
            raise

        self.context = self.listener.CreateRecoContext()
        self.grammar = self.context.CreateGrammar()
        self.grammar.DictationSetState(0)

        self.wordsRule = self.grammar.Rules.Add("wordsRule", constants.SRATopLevel | constants.SRADynamic, 0)
        self.wordsRule.Clear()
        [self.wordsRule.InitialState.AddWordTransition(None, word) for word in wordsToAdd]
        
        self.grammar.Rules.Commit()
        self.grammar.CmdSetRuleState("wordsRule", 1)
        self.grammar.Rules.Commit()
        
        self.eventHandler = ContextEvents(self.context, self.speaker, self.config_data, self.grammar)
        
        print("后台语音识别已成功启动，正在监听...")


    def _set_voice(self, voice_name):
        voices = self.speaker.GetVoices()
        found_voice = None
        for voice in voices:
            if voice_name.lower() in voice.GetDescription().lower():
                found_voice = voice
                break
        if found_voice:
            self.speaker.Voice = found_voice
            print(f"语音已成功设置为: {self.speaker.Voice.GetDescription()}")
        else:
            print(f"警告: 未找到名为 '{voice_name}' 的语音，将使用默认语音。")

    def _set_recognizer(self, recognizer_name):
        cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
        cat.SetId(constants.SpeechCategoryRecognizers)
        recognizers = cat.EnumerateTokens()
        found_reco = None
        for reco in recognizers:
            if recognizer_name.lower() in reco.GetDescription().lower():
                found_reco = reco
                break
        if found_reco:
            self.listener.Recognizer = found_reco
            print(f"识别器已成功设置为: {self.listener.Recognizer.GetDescription()}")
        else:
            print(f"警告: 未找到名为 '{recognizer_name}' 的识别器，将使用默认识别器。")

    def say(self, phrase):
        self.speaker.Speak(phrase)


class ContextEvents(win32com.client.getevents("SAPI.SpSharedRecoContext")):
    def __init__(self, context, speaker, config_data, grammar):
        super().__init__(context)
        self.speaker = speaker
        self.grammar = grammar
        self.command_map = {cmd['trigger']: cmd for cmd in config_data.get('commands', [])}
        self.is_executing = False

    def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result):
        if self.is_executing:
            print("  -> 正在执行上一条指令，忽略新的识别结果。")
            return
            
        newResult = win32com.client.Dispatch(Result)
        recognized_text = newResult.PhraseInfo.GetText()
        
        print(f"\n识别到指令: '{recognized_text}'")
        
        command_data = self.command_map.get(recognized_text)
        
        if command_data:
            self.is_executing = True
            action_thread = threading.Thread(target=self._run_actions, args=(command_data['actions'],))
            action_thread.start()
        else:
            print("  -> 该指令未在配置文件中定义，无响应动作。")

    def _run_actions(self, actions):
        try:
            print("  -> 识别已暂停...")
            self.grammar.CmdSetRuleState("wordsRule", 0) # 0 = Inactive
            self.grammar.Rules.Commit()

            for action in actions:
                action_type = action.get('type')
                params = action.get('params', {})
                print(f"  -> 执行动作: {action_type}, 参数: {params}")
                
                if action_type == 'speak':
                    text = params.get('text')
                    if text:
                        flags = constants.SVSFlagsAsync | constants.SVSFDefault
                        self.speaker.Speak(text, flags)
                elif action_type == 'keypress':
                    self._execute_keypress(params)
                elif action_type == 'write':
                    self._execute_write(params)
                
                if "delay_after" in action:
                    time.sleep(action["delay_after"])

            self.speaker.WaitUntilDone(10000) 

        except Exception as e:
            print(f"  -> 执行动作时出错: {e}")
        finally:
            self.grammar.CmdSetRuleState("wordsRule", 1) # 1 = Active
            self.grammar.Rules.Commit()
            self.is_executing = False
            print("  -> 识别已恢复。")

    def _execute_keypress(self, params):
        keys = params.get('keys')
        if keys and isinstance(keys, list):
            print(f"    -> 模拟游戏按键 (Manual): {keys}")
            try:
                # pydirectinput 没有 hotkey 函数，我们需要手动按顺序按下
                for key in keys:
                    pydirectinput.keyDown(key)
                    # 稍微延迟一下，防止游戏检测不到修饰键（如Ctrl）的按下
                    time.sleep(0.05) 
                
                # ★ 关键：保持按键状态一小段时间（0.1~0.2秒）
                # 很多游戏如果按键时间短于一帧，会忽略输入
                time.sleep(0.1)

                # 反向释放按键（先松开后按下的键）
                for key in reversed(keys):
                    pydirectinput.keyUp(key)
                    # 释放时也稍微给点缓冲，避免卡键
                    time.sleep(0.05)
                    
            except Exception as e:
                print(f"    -> 按键模拟失败: {e}")

    def _execute_write(self, params):
        text = params.get('text')
        if text is not None:
            print(f"    -> 输入文本: {text}")
            # 方法 A: 尝试直接使用 write (有些版本可能叫 typewrite)
            if hasattr(pydirectinput, 'write'):
                pydirectinput.write(text, interval=0.1)
            elif hasattr(pydirectinput, 'typewrite'):
                pydirectinput.typewrite(text, interval=0.1)
            else:
                # 方法 B: 如果都没有，就逐个按键
                for char in text:
                    pydirectinput.press(char)
                    time.sleep(0.1)

def load_config(file_path='config.json'):
    """
    加载并验证JSON配置文件。
    在 PyInstaller 打包后，文件可能在临时目录中。
    """
    if getattr(sys, 'frozen', False):
        # 如果是 PyInstaller 打包的程序
        # sys.executable 是当前运行的 EXE 文件的完整路径
        base_path = os.path.dirname(sys.executable)
        # 注意：sys._MEIPASS 是解压临时目录，不是EXE所在目录，用于访问打包在EXE内部的文件。
        # 如果config.json是外部文件，则不应使用sys._MEIPASS。
    else:
        # 如果是普通Python脚本，文件路径是脚本所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))

    absolute_path = os.path.join(base_path, file_path)

    print(f"尝试加载配置文件: {absolute_path}") # 添加打印，方便调试

    if not os.path.exists(absolute_path):
        raise FileNotFoundError(f"配置文件 '{file_path}' 未找到。请确保它与 EXE 文件（或脚本）在同一目录: '{base_path}'。")
    
    with open(absolute_path, 'r', encoding='utf-8') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"配置文件 '{absolute_path}' 格式错误: {e}")

if __name__ == '__main__':
    # 首次运行时，仍然建议先运行此函数来查看可用模型
    # list_sapi_resources() 

    try:
        config_data = load_config('./config.json')
        
        # 从配置文件中获取 settings
        settings = config_data.get('settings', {})
        voice_name = settings.get('voice_name')
        recognizer_name = settings.get('recognizer_name')
        speech_rate = settings.get('speech_rate', 0) # 默认语速为0
        initial_greeting = settings.get('initial_greeting', "")

        # 使用从 config.json 读取的设置来初始化 SpeechRecognition
        speechReco = SpeechRecognition(config_data, 
                                       voice_name=voice_name, 
                                       recognizer_name=recognizer_name, 
                                       speech_rate=speech_rate,
                                       initial_greeting=initial_greeting) 
        
        print("\n--- Initialization complete, main loop starting ---")
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)

    except (KeyboardInterrupt, SystemExit):
        print("\nProgram interrupted or exited by user...")
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"\nProgram failed to start: {e}")
    except Exception as e:
        print(f"\nAn unexpected error occurred during runtime: {e}")
    finally:
        print("Program ended.")