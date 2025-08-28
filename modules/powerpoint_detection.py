"""
PowerPoint COM API を使用した動画再生検出モジュール
CPU使用率ベースの判定の代替として、より堅牢な動画再生検出を提供
"""
import logging
import traceback

try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError as e:
    WIN32_AVAILABLE = False
    print(f"PowerPoint COM検出が無効: {e}")


class PowerPointDetector:
    """PowerPoint COM APIを使用した動画再生検出クラス"""

    def __init__(self, debug_mode=False):
        self.debug_mode = debug_mode
        self.app = None
        self.last_error = None

    def _log_debug(self, message):
        """デバッグログ出力"""
        if self.debug_mode:
            print(f"[PowerPoint検出] {message}")

    def _log_error(self, message):
        """エラーログ出力"""
        print(f"[PowerPoint検出エラー] {message}")
        self.last_error = message

    def _get_powerpoint_app(self):
        """PowerPointアプリケーションのCOMオブジェクトを取得"""
        try:
            if not WIN32_AVAILABLE:
                return None

            # 既存のPowerPointインスタンスに接続を試行
            self.app = win32com.client.GetActiveObject(
                "PowerPoint.Application")
            self._log_debug("既存のPowerPointアプリケーションに接続しました")
            return self.app

        except Exception as e:
            # 既存のインスタンスがない場合やエラーの場合
            self._log_debug(f"PowerPointアプリケーション取得エラー: {e}")
            self.app = None
            return None

    def _check_slideshow_window(self):
        """スライドショーウィンドウの存在とフルスクリーン状態を確認"""
        try:
            if not self.app:
                return False

            # スライドショーウィンドウの数を確認
            slideshow_count = self.app.SlideShowWindows.Count
            self._log_debug(f"スライドショーウィンドウ数: {slideshow_count}")

            if slideshow_count == 0:
                return False

            # フォアグラウンドウィンドウがPowerPointのスライドショーかどうかを確認
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd).lower()
            class_name = win32gui.GetClassName(hwnd).lower()

            # PowerPointのスライドショーウィンドウクラス
            slideshow_classes = ['pptframeclass', 'screenclass']

            is_powerpoint_window = ('powerpoint' in title or
                                    any(cls in class_name for cls in slideshow_classes))

            if not is_powerpoint_window:
                self._log_debug("フォアグラウンドウィンドウがPowerPointスライドショーではありません")
                return False

            # フルスクリーン判定
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]

            screen_width = win32con.SM_CXSCREEN
            screen_height = win32con.SM_CYSCREEN

            # システムメトリクスを取得
            try:
                import win32api
                screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
                screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            except:
                # フォールバック値
                screen_width = 1920
                screen_height = 1080

            is_fullscreen = (width >= screen_width -
                             20 and height >= screen_height - 20)

            self._log_debug(
                f"ウィンドウサイズ: {width}x{height}, スクリーン: {screen_width}x{screen_height}, フルスクリーン: {is_fullscreen}")

            return is_fullscreen

        except Exception as e:
            self._log_error(f"スライドショーウィンドウ確認エラー: {e}")
            return False

    def _check_media_in_current_slide(self):
        """現在のスライドに動画があるかどうかを確認（改良版：サイズに関係なく検出）"""
        try:
            if not self.app:
                return False

            slideshow_windows = self.app.SlideShowWindows
            if slideshow_windows.Count == 0:
                return False

            # 最初のスライドショーウィンドウを取得
            slideshow_window = slideshow_windows(1)
            current_slide = slideshow_window.View.Slide

            self._log_debug(f"現在のスライド番号: {current_slide.SlideIndex}")

            # スライド内のシェイプを確認
            shapes = current_slide.Shapes
            media_shapes = []

            for shape in shapes:
                try:
                    # Type 16 はメディア（動画/音声）オブジェクト
                    if shape.Type == 16:  # msoMedia
                        # サイズ情報も取得して記録
                        try:
                            width = shape.Width
                            height = shape.Height
                            area = width * height
                            self._log_debug(
                                f"メディアオブジェクト発見: {shape.Name} (サイズ: {width:.1f}x{height:.1f}, 面積: {area:.0f})")

                            # サイズに関係なくメディアオブジェクトとして追加
                            media_shapes.append(shape)

                            # 小さな動画の場合は特別にログ出力
                            if area < 10000:  # 面積が10000平方ポイント未満を小さな動画とする
                                self._log_debug(
                                    f"小さな動画検出: {shape.Name} (面積: {area:.0f})")
                        except:
                            # サイズ取得に失敗してもメディアオブジェクトとして処理
                            media_shapes.append(shape)
                            self._log_debug(
                                f"メディアオブジェクト発見: {shape.Name} (サイズ情報取得不可)")

                except Exception as shape_error:
                    # 一部のシェイプでエラーが発生してもスキップ
                    continue

            has_media = len(media_shapes) > 0
            self._log_debug(f"メディアシェイプ数: {len(media_shapes)}")

            return has_media, media_shapes

        except Exception as e:
            self._log_error(f"スライド内メディア確認エラー: {e}")
            return False, []

    def _check_media_playback_state(self, media_shapes):
        """メディアオブジェクトの再生状態を確認（改良版）"""
        try:
            playing_media = []
            state_unknown_media = []

            for shape in media_shapes:
                media_info = {
                    'name': shape.Name,
                    'shape': shape,
                    'play_state': None,
                    'play_state_name': 'Unknown',
                    'access_method': None,
                    'additional_info': {}
                }

                # 方法1: 標準的なMediaFormat.PlayStateアクセス
                play_state_found = self._try_standard_playstate(
                    shape, media_info)

                # 方法2: 代替アクセス方法
                if not play_state_found:
                    play_state_found = self._try_alternative_playstate(
                        shape, media_info)

                # 方法3: Animation効果からの推測
                if not play_state_found:
                    play_state_found = self._try_animation_state(
                        shape, media_info)

                # 結果の評価
                if media_info['play_state'] == 1:  # 再生中
                    playing_media.append(shape)
                    self._log_debug(
                        f"再生中のメディア発見: {shape.Name} (方法: {media_info['access_method']})")
                elif media_info['play_state'] is not None:
                    state_name = self._get_playstate_name(
                        media_info['play_state'])
                    self._log_debug(
                        f"メディア {shape.Name}: {state_name} (方法: {media_info['access_method']})")
                else:
                    state_unknown_media.append(shape)
                    self._log_debug(f"メディア {shape.Name}: 再生状態不明 (すべての方法が失敗)")

            # 結果のサマリー
            total_media = len(media_shapes)
            playing_count = len(playing_media)
            unknown_count = len(state_unknown_media)

            self._log_debug(
                f"メディア状態サマリー: 総数={total_media}, 再生中={playing_count}, 不明={unknown_count}")

            # 再生中のメディアがある場合は確実に抑制
            if playing_count > 0:
                return True, playing_media

            # 不明なメディアがある場合の判定（改良版：より保守的）
            if unknown_count > 0:
                # 不明なメディアがある場合は、安全のためFalseを返す（抑制しない）
                # 理由：確実に再生中と判定できない限り、スクリーンセーバーを無効にすべきではない
                self._log_debug(
                    f"{unknown_count}個のメディアの再生状態が不明です。確実性がないため抑制しません")
                return False, state_unknown_media

            # すべてのメディアが停止中/一時停止中
            return False, []

        except Exception as e:
            self._log_error(f"メディア再生状態確認エラー: {e}")
            if self.debug_mode:
                traceback.print_exc()
            return False, []

    def _try_standard_playstate(self, shape, media_info):
        """標準的なMediaFormat.PlayStateアクセスを試行（改良版）"""
        try:
            if hasattr(shape, 'MediaFormat'):
                media_format = shape.MediaFormat

                # PlayStateの直接取得を試行
                if hasattr(media_format, 'PlayState'):
                    try:
                        play_state = media_format.PlayState
                        media_info['play_state'] = play_state
                        media_info['play_state_name'] = self._get_playstate_name(
                            play_state)
                        media_info['access_method'] = 'MediaFormat.PlayState'
                        return True
                    except Exception as e:
                        self._log_debug(f"PlayState直接取得失敗 ({shape.Name}): {e}")

                # PlayStateが取得できない場合、他のプロパティから推測
                return self._infer_playstate_from_properties(shape, media_info, media_format)

        except Exception as e:
            self._log_debug(f"標準PlayState取得エラー ({shape.Name}): {e}")

        return False

    def _infer_playstate_from_properties(self, shape, media_info, media_format):
        """MediaFormatの他のプロパティから再生状態を推測"""
        try:
            # additional_infoの初期化を確認
            if 'additional_info' not in media_info:
                media_info['additional_info'] = {}

            # 利用可能なプロパティを収集
            properties = {}

            # 基本プロパティの取得
            for prop_name in ['Length', 'Volume', 'Muted', 'StartPoint', 'EndPoint']:
                try:
                    value = getattr(media_format, prop_name)
                    properties[prop_name] = value
                    media_info['additional_info'][prop_name] = value
                except:
                    pass

            # ActionSettingsから推測
            action_info = self._get_action_settings_info(shape)
            if action_info:
                properties.update(action_info)
                media_info['additional_info'].update(action_info)

            # アニメーション効果から推測
            animation_info = self._get_animation_info(shape)
            if animation_info:
                properties.update(animation_info)
                media_info['additional_info'].update(animation_info)

            # 推測ロジック
            inferred_state = self._infer_state_from_properties(
                properties, shape.Name)

            if inferred_state is not None:
                media_info['play_state'] = inferred_state
                media_info['play_state_name'] = self._get_playstate_name(
                    inferred_state)
                media_info['access_method'] = 'プロパティ推測'
                self._log_debug(
                    f"プロパティから再生状態を推測: {shape.Name} -> {media_info['play_state_name']}")
                return True

        except Exception as e:
            self._log_debug(f"プロパティ推測エラー ({shape.Name}): {e}")

        return False

    def _get_action_settings_info(self, shape):
        """ActionSettingsから情報を取得"""
        action_info = {}
        try:
            if hasattr(shape, 'ActionSettings'):
                # MouseClick (1) の情報を取得
                action_setting = shape.ActionSettings(1)
                if hasattr(action_setting, 'Action'):
                    action = action_setting.Action
                    action_info['click_action'] = action

                    # アクション12は特定の動作（再生関連の可能性）
                    if action == 12:
                        action_info['has_play_action'] = True
        except:
            pass
        return action_info

    def _get_animation_info(self, shape):
        """アニメーション効果から情報を取得"""
        animation_info = {}
        try:
            slide = shape.Parent
            if hasattr(slide, 'TimeLine') and hasattr(slide.TimeLine, 'MainSequence'):
                main_sequence = slide.TimeLine.MainSequence
                for i in range(1, main_sequence.Count + 1):
                    effect = main_sequence.Item(i)
                    if hasattr(effect, 'Shape') and effect.Shape.Name == shape.Name:
                        if hasattr(effect, 'EffectType'):
                            effect_type = effect.EffectType
                            animation_info['effect_type'] = effect_type

                            # エフェクトタイプ83は何らかのメディア関連効果
                            if effect_type == 83:
                                animation_info['has_media_effect'] = True

                        if hasattr(effect, 'Timing') and hasattr(effect.Timing, 'TriggerType'):
                            trigger_type = effect.Timing.TriggerType
                            animation_info['trigger_type'] = trigger_type
                        break
        except:
            pass
        return animation_info

    def _infer_state_from_properties(self, properties, shape_name):
        """収集したプロパティから再生状態を推測（改良版：小さな動画対応）"""

        # 音量が0でミュートされていない場合、再生中の可能性が高い
        volume = properties.get('Volume', 0)
        muted = properties.get('Muted', True)

        # アクション設定の確認
        has_play_action = properties.get('has_play_action', False)
        has_media_effect = properties.get('has_media_effect', False)

        # Length, StartPoint, EndPointの確認（再生中判定のヒント）
        length = properties.get('Length', 0)
        start_point = properties.get('StartPoint', 0)
        end_point = properties.get('EndPoint', 0)

        # 小さな動画用の追加判定要素
        click_action = properties.get('click_action', 0)
        trigger_type = properties.get('trigger_type', 0)
        effect_type = properties.get('effect_type', 0)

        self._log_debug(
            f"推測材料 ({shape_name}): Volume={volume}, Muted={muted}, PlayAction={has_play_action}, MediaEffect={has_media_effect}, Length={length}, ClickAction={click_action}")

        # 小さな動画の場合、より積極的な判定を行う
        # 小さな動画では一部のプロパティが正しく取得できない場合がある

        # 1. 確実に再生中と判定できる条件
        if has_play_action and has_media_effect and length > 0:
            if volume > 0 and not muted:
                # 通常サイズと同じ条件だが、小さな動画でも適用
                self._log_debug(f"推測結果: 再生中の可能性高（音量・アクション条件満たす）")
                return 1  # 再生中
            elif volume > 0 or not muted:
                # 片方の条件のみ満たす場合でも再生中の可能性
                self._log_debug(f"推測結果: 再生中の可能性（部分的音量条件）")
                return 1  # 再生中

        # 2. アクション設定のみでも判定（小さな動画対応）
        if has_play_action or (click_action == 12):  # アクション12は再生関連
            if length > 0:
                self._log_debug(f"推測結果: 状態不明（アクション設定あり、確実性なし）")
                return None  # 不明（保守的判定）

        # 3. アニメーション効果からの判定（小さな動画でも有効）
        if has_media_effect or (effect_type == 83):  # 効果タイプ83はメディア関連
            if length > 0:
                self._log_debug(f"推測結果: 状態不明（アニメーション効果あり、確実性なし）")
                return None  # 不明（保守的判定）

        # 4. 基本的な停止判定
        if length > 0:
            if muted and volume == 0:
                self._log_debug(f"推測結果: 停止中（ミュート・音量0）")
                return 2  # 停止
            else:
                # メディアファイルは存在するが状態が不明
                self._log_debug(f"推測結果: 状態不明（メディア存在、条件不明）")
                return None

        # 5. デフォルト：停止中
        self._log_debug(f"推測結果: 停止中（デフォルト）")
        return 2  # 停止

    def _try_alternative_playstate(self, shape, media_info):
        """代替的なアクセス方法を試行（小さな動画対応強化）"""
        try:
            # 方法2-1: ActionSettingsからの推測（詳細化）
            if hasattr(shape, 'ActionSettings'):
                try:
                    action_settings = shape.ActionSettings(1)  # ppMouseClick
                    if hasattr(action_settings, 'Action'):
                        action = action_settings.Action
                        media_info['additional_info']['click_action'] = action

                        if action == 32:  # ppActionPlay
                            media_info['access_method'] = 'ActionSettings推測(Play)'
                            self._log_debug(f"アクション設定でPlayが検出: {shape.Name}")
                        elif action == 12:  # 特定のアクション（再生関連の可能性）
                            media_info['access_method'] = 'ActionSettings推測(Action12)'
                            self._log_debug(f"アクション12が検出: {shape.Name}")

                        # 他のActionSettings項目も確認
                        if hasattr(action_settings, 'AnimateAction'):
                            animate_action = action_settings.AnimateAction
                            media_info['additional_info']['animate_action'] = animate_action

                except Exception as e:
                    self._log_debug(f"ActionSettings取得エラー ({shape.Name}): {e}")

            # 方法2-2: OLEObjectからの情報取得（拡張）
            if hasattr(shape, 'OLEFormat'):
                try:
                    ole_format = shape.OLEFormat
                    if hasattr(ole_format, 'Object'):
                        ole_object = ole_format.Object

                        # より多くのプロパティをチェック
                        ole_properties = ['PlayState',
                                          'CurrentState', 'State', 'Status']
                        for prop_name in ole_properties:
                            try:
                                if hasattr(ole_object, prop_name):
                                    prop_value = getattr(ole_object, prop_name)
                                    media_info['play_state'] = prop_value
                                    media_info['access_method'] = f'OLEFormat.Object.{prop_name}'
                                    self._log_debug(
                                        f"OLE {prop_name}から状態取得: {shape.Name} = {prop_value}")
                                    return True
                            except:
                                continue

                        # ProgIDの確認（小さな動画の種類判定）
                        try:
                            if hasattr(ole_format, 'ProgID'):
                                prog_id = ole_format.ProgID
                                media_info['additional_info']['prog_id'] = prog_id
                                self._log_debug(
                                    f"ProgID: {shape.Name} = {prog_id}")
                        except:
                            pass

                except Exception as e:
                    self._log_debug(f"OLEFormat取得エラー ({shape.Name}): {e}")

            # 方法2-3: AnimationEffectsからの推測（詳細化）
            try:
                slide = shape.Parent
                if hasattr(slide, 'TimeLine') and hasattr(slide.TimeLine, 'MainSequence'):
                    main_sequence = slide.TimeLine.MainSequence

                    # すべてのエフェクトをチェック
                    for i in range(1, main_sequence.Count + 1):
                        try:
                            effect = main_sequence.Item(i)
                            if hasattr(effect, 'Shape') and effect.Shape.Name == shape.Name:

                                # エフェクトタイプの詳細チェック
                                if hasattr(effect, 'EffectType'):
                                    effect_type = effect.EffectType
                                    media_info['additional_info']['effect_type'] = effect_type

                                    # 既知のメディア関連エフェクト
                                    # 各種メディア再生エフェクト
                                    media_effects = [30, 83, 84, 85]
                                    if effect_type in media_effects:
                                        media_info['access_method'] = f'Animation効果推測(Type{effect_type})'
                                        self._log_debug(
                                            f"メディア関連アニメーション効果検出: {shape.Name} (Type: {effect_type})")

                                # タイミング情報の取得
                                if hasattr(effect, 'Timing'):
                                    timing = effect.Timing
                                    if hasattr(timing, 'TriggerType'):
                                        trigger_type = timing.TriggerType
                                        media_info['additional_info']['trigger_type'] = trigger_type

                                    # 再生時間の情報
                                    if hasattr(timing, 'Duration'):
                                        duration = timing.Duration
                                        media_info['additional_info']['effect_duration'] = duration

                                break
                        except:
                            continue

            except Exception as e:
                self._log_debug(f"アニメーション効果取得エラー ({shape.Name}): {e}")

            # 方法2-4: 小さな動画特有の検出（新規追加）
            try:
                # 動画のファイル拡張子から種類を判定
                if hasattr(shape, 'MediaFormat') and hasattr(shape.MediaFormat, 'FileName'):
                    filename = shape.MediaFormat.FileName
                    media_info['additional_info']['filename'] = filename

                    # 小さな動画でよく使われる形式
                    small_video_formats = ['.gif', '.webm', '.mp4']
                    if any(filename.lower().endswith(fmt) for fmt in small_video_formats):
                        self._log_debug(
                            f"小さな動画形式検出: {shape.Name} ({filename})")
                        media_info['additional_info']['small_video_format'] = True

            except Exception as e:
                self._log_debug(f"ファイル名取得エラー ({shape.Name}): {e}")

        except Exception as e:
            self._log_debug(f"代替PlayState取得エラー ({shape.Name}): {e}")

        return False

    def _try_animation_state(self, shape, media_info):
        """アニメーション効果からの状態推測"""
        try:
            slide = shape.Parent
            if hasattr(slide, 'SlideShowTransition'):
                # スライドショー中のアニメーション状態を確認
                pass  # 詳細な実装は必要に応じて追加
        except:
            pass

        return False

    def _get_playstate_name(self, play_state):
        """PlayState値を人間読み可能な名前に変換"""
        state_names = {
            0: "一時停止",
            1: "再生中",
            2: "停止",
            3: "準備中",
            4: "バッファリング中",
            5: "エラー",
            6: "終了",
            -1: "不明"
        }
        return state_names.get(play_state, f"未定義({play_state})")

    def is_video_playing_in_slideshow(self):
        """PowerPointスライドショーで動画が再生中かどうかを判定（改良版）"""
        try:
            # COM初期化
            try:
                pythoncom.CoInitialize()
            except:
                pass  # 既に初期化済みの場合はスキップ

            # PowerPointアプリケーションを取得
            if not self._get_powerpoint_app():
                self._log_debug("PowerPointアプリケーションが見つかりません")
                return False

            # スライドショーウィンドウの確認
            if not self._check_slideshow_window():
                self._log_debug("PowerPointスライドショーが実行中でないか、フルスクリーンではありません")
                return False

            # 現在のスライドでメディアの有無を確認
            has_media_result = self._check_media_in_current_slide()
            if isinstance(has_media_result, tuple):
                has_media, media_shapes = has_media_result
            else:
                has_media = has_media_result
                media_shapes = []

            if not has_media:
                self._log_debug("現在のスライドにメディアオブジェクトがありません")
                return False

            # メディアの再生状態を確認（改良版）
            is_playing, playing_media = self._check_media_playback_state(
                media_shapes)

            if is_playing:
                # 確実に再生中のメディアが検出された場合
                if len(playing_media) > 0 and hasattr(playing_media[0], 'Name'):
                    media_names = [getattr(media, 'Name', 'Unknown')
                                   for media in playing_media]
                    self._log_debug(
                        f"動画再生中確認 - 再生中メディア: {', '.join(media_names)}")
                else:
                    self._log_debug("動画再生中確認 - 詳細不明ですが再生中と判定")
                return True
            else:
                # 再生中ではない、または不明な場合は抑制しない
                self._log_debug("メディアオブジェクトは存在しますが、再生中ではないか不明のため抑制しません")
                return False

        except Exception as e:
            self._log_error(f"PowerPoint動画再生判定エラー: {e}")
            if self.debug_mode:
                traceback.print_exc()
            return False

        finally:
            # COM終了処理
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def get_powerpoint_info(self):
        """PowerPoint の詳細情報を取得（デバッグ用・改良版）"""
        try:
            pythoncom.CoInitialize()

            if not self._get_powerpoint_app():
                return {"error": "PowerPointアプリケーションが見つかりません"}

            info = {
                "slideshow_count": self.app.SlideShowWindows.Count,
                "presentations_count": self.app.Presentations.Count,
                "version": getattr(self.app, 'Version', 'Unknown'),
                "current_slide": None,
                "media_objects": []
            }

            if info["slideshow_count"] > 0:
                slideshow = self.app.SlideShowWindows(1)
                current_slide = slideshow.View.Slide
                info["current_slide"] = {
                    "index": current_slide.SlideIndex,
                    "name": getattr(current_slide, 'Name', 'Unknown'),
                    "shapes_count": current_slide.Shapes.Count
                }

                # メディアオブジェクトの詳細（改良版）
                for shape in current_slide.Shapes:
                    try:
                        if shape.Type == 16:  # メディア
                            media_info = {
                                "name": shape.Name,
                                "type": "Media",
                                "play_state": None,
                                "play_state_name": "Unknown",
                                "access_method": None,
                                "properties": {}
                            }

                            # 改良された再生状態チェック
                            self._try_standard_playstate(shape, media_info)
                            if media_info['play_state'] is None:
                                self._try_alternative_playstate(
                                    shape, media_info)

                            # 基本プロパティの取得
                            try:
                                if hasattr(shape, 'MediaFormat'):
                                    media_format = shape.MediaFormat
                                    if hasattr(media_format, 'FileName'):
                                        media_info["properties"]["filename"] = media_format.FileName
                                    if hasattr(media_format, 'Length'):
                                        media_info["properties"]["length"] = media_format.Length
                                    if hasattr(media_format, 'Position'):
                                        media_info["properties"]["position"] = media_format.Position
                                    if hasattr(media_format, 'Volume'):
                                        media_info["properties"]["volume"] = media_format.Volume
                                    if hasattr(media_format, 'Muted'):
                                        media_info["properties"]["muted"] = media_format.Muted
                            except:
                                pass

                            info["media_objects"].append(media_info)
                    except:
                        continue

            return info

        except Exception as e:
            return {"error": str(e)}
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def get_detailed_media_state(self):
        """メディアオブジェクトの詳細な状態情報を取得"""
        try:
            pythoncom.CoInitialize()

            if not self._get_powerpoint_app():
                return {"error": "PowerPointアプリケーションが見つかりません"}

            if self.app.SlideShowWindows.Count == 0:
                return {"error": "スライドショーが実行中ではありません"}

            slideshow = self.app.SlideShowWindows(1)
            current_slide = slideshow.View.Slide

            detailed_info = {
                "slide_index": current_slide.SlideIndex,
                "media_objects": [],
                "summary": {
                    "total_media": 0,
                    "playing": 0,
                    "paused": 0,
                    "stopped": 0,
                    "unknown": 0
                }
            }

            for shape in current_slide.Shapes:
                if shape.Type == 16:  # メディア
                    media_info = {
                        'name': shape.Name,
                        'play_state': None,
                        'play_state_name': 'Unknown',
                        'access_method': None,
                        'additional_info': {}
                    }

                    # 詳細な状態チェック
                    self._try_standard_playstate(shape, media_info)
                    if media_info['play_state'] is None:
                        self._try_alternative_playstate(shape, media_info)

                    detailed_info["media_objects"].append(media_info)
                    detailed_info["summary"]["total_media"] += 1

                    # 統計の更新
                    if media_info['play_state'] == 1:
                        detailed_info["summary"]["playing"] += 1
                    elif media_info['play_state'] == 0:
                        detailed_info["summary"]["paused"] += 1
                    elif media_info['play_state'] == 2:
                        detailed_info["summary"]["stopped"] += 1
                    else:
                        detailed_info["summary"]["unknown"] += 1

            return detailed_info

        except Exception as e:
            return {"error": str(e)}
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


# シングルトンインスタンス
_powerpoint_detector = None


def get_powerpoint_detector(debug_mode=False):
    """PowerPointDetectorのシングルトンインスタンスを取得"""
    global _powerpoint_detector
    if _powerpoint_detector is None:
        _powerpoint_detector = PowerPointDetector(debug_mode=debug_mode)
    return _powerpoint_detector


def is_powerpoint_video_playing(debug_mode=False):
    """PowerPointで動画が再生中かどうかを判定（外部から使用する関数）"""
    if not WIN32_AVAILABLE:
        return False

    detector = get_powerpoint_detector(debug_mode=debug_mode)
    return detector.is_video_playing_in_slideshow()


def get_powerpoint_debug_info():
    """PowerPointのデバッグ情報を取得"""
    if not WIN32_AVAILABLE:
        return {"error": "win32com not available"}

    detector = get_powerpoint_detector(debug_mode=True)
    return detector.get_powerpoint_info()


def get_detailed_powerpoint_media_state():
    """PowerPointメディアの詳細状態を取得"""
    if not WIN32_AVAILABLE:
        return {"error": "win32com not available"}

    detector = get_powerpoint_detector(debug_mode=True)
    return detector.get_detailed_media_state()


if __name__ == "__main__":
    # テスト実行
    print("PowerPoint COM API 動画検出テスト")
    print("=" * 50)

    # デバッグ情報を表示
    info = get_powerpoint_debug_info()
    print("PowerPoint情報:")
    for key, value in info.items():
        print(f"  {key}: {value}")

    print("\n動画再生状態:")
    is_playing = is_powerpoint_video_playing(debug_mode=True)
    print(f"動画再生中: {'はい' if is_playing else 'いいえ'}")
