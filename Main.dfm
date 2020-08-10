object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'SkySlope Ripper'
  ClientHeight = 596
  ClientWidth = 1595
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object WebBrowser2: TWebBrowser
    Left = 1151
    Top = 39
    Width = 46
    Height = 17
    TabOrder = 2
    OnBeforeNavigate2 = WebBrowser2BeforeNavigate2
    ControlData = {
      4C000000C1040000C20100000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Browse'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 89
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Grab'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Memo1: TMemo
    Left = 1203
    Top = 39
    Width = 382
    Height = 386
    TabOrder = 3
  end
  object Button3: TButton
    Left = 170
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Jump To Page'
    Enabled = False
    TabOrder = 4
    OnClick = Button3Click
  end
  object Memo2: TMemo
    Left = 1203
    Top = 431
    Width = 384
    Height = 154
    TabOrder = 5
  end
  object WebBrowser1: TWebBrowser
    Left = 8
    Top = 39
    Width = 1189
    Height = 546
    TabOrder = 6
    OnNewWindow2 = WebBrowser1NewWindow2
    ControlData = {
      4C000000E37A00006E3800000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126208000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object TimerNextPage: TTimer
    Enabled = False
    Interval = 3000
    OnTimer = TimerNextPageTimer
    Left = 104
    Top = 96
  end
  object MemoSaveTimer: TTimer
    Interval = 10000
    OnTimer = MemoSaveTimerTimer
    Left = 1208
    Top = 48
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 3000
    OnTimer = Timer1Timer
    Left = 192
    Top = 40
  end
  object TimeoutTimer: TTimer
    Enabled = False
    Interval = 240000
    OnTimer = TimeoutTimerTimer
    Left = 104
    Top = 40
  end
end
