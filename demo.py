# -*- coding: utf-8 -*-

import win32com.client as win32
import pywintypes
import csv
import sys

fontPoints = {'八号': 5,
              '七号': 5.5,
              '小六': 6.5,
              '六号': 7.5,
              '小五': 9,
              '五号': 10.5,
              '小四': 12,
              '四号': 14,
              '小三': 15,
              '三号': 16,
              '小二': 18,
              '二号': 22,
              '小一': 24,
              '一号': 26,
              '小初': 36,
              '初号': 42
              }
keyCode = {'0': 48,
           '1': 49,
           '2': 50,
           '3': 51,
           '4': 52,
           '5': 53,
           '6': 54,
           '7': 55,
           '8': 56,
           '9': 57,
           'a': 65,
           'b': 66,
           'c': 67,
           'd': 68,
           'e': 69,
           'f': 70,
           'g': 71,
           'h': 72,
           'i': 73,
           'j': 74,
           'k': 75,
           'l': 76,
           'm': 77,
           'n': 78,
           'o': 79,
           'p': 80,
           'q': 81,
           'r': 82,
           's': 83,
           't': 84,
           'u': 85,
           'v': 86,
           'w': 87,
           'x': 88,
           'y': 89,
           'z': 90,
           'ctrl': 512,
           'alt': 1024,
           'shift': 256,
           'esc': 27,
           '`': 192,
           '/': 111,
           '[\]': 220,
           'tab': 9,
           'f1': 112,
           'f2': 113,
           'f3': 114,
           'f4': 115,
           'f5': 116,
           'f6': 117,
           'f7': 118,
           'f8': 119,
           'f9': 120,
           'f10': 121,
           'f11': 122,
           'f12': 123,
           }

def loadStylesFromCsv(file):
    app = win32.gencache.EnsureDispatch('Word.Application')
    try:
        doc = app.ActiveDocument
    except pywintypes.com_error as e:
        print('com_error:', e)
        sys.exit(1)
    with open(file, 'r', newline='') as f:
        dictReader = csv.DictReader(f)
        for style in dictReader:
            if style['Bold'] == 'FALSE' or style['Bold'] == 'False' or style['Bold'] == 'false':
                style['Bold'] = False
            elif style['Bold'] == 'TRUE' or style['Bold'] == 'True' or style['Bold'] == 'true':
                style['Bold'] = True
            if style['Italic'] == 'FALSE' or style['Italic'] == 'False' or style['Italic'] == 'false':
                style['Italic'] = False
            elif style['Italic'] == 'TRUE' or style['Italic'] == 'True' or style['Italic'] == 'true':
                style['Italic'] = True
            style['OutlineLevel'] = int(style['OutlineLevel'])
            style['LeftIndent'] = int(style['LeftIndent'])
            style['RightIndent'] = int(style['RightIndent'])
            style['FirstLineIndent'] = int(style['FirstLineIndent'])
            style['LineUnitBefore'] = int(style['LineUnitBefore'])
            style['LineUnitAfter'] = int(style['LineUnitAfter'])
            style['LineSpacing'] = float(style['LineSpacing'])
            newStyle(app, doc, style)

# def linesToPoints(line):
#     return 12*line

def newStyle(app, doc, fmt):
    try:
        doc.Styles(fmt['Name'])
    except pywintypes.com_error as e:
        if e.hresult == -2147352567:
            doc.Styles.Add(fmt['Name'], 1)
        else:
            print('com_error:', e)
            sys.exit(1)
    doc.Styles(fmt['Name']).AutomaticallyUpdate = False
    if fmt['Name'] != '正文':
        doc.Styles(fmt['Name']).BaseStyle = "正文"
    doc.Styles(fmt['Name']).NextParagraphStyle = "正文"
    doc.Styles(fmt['Name']).Font.NameFarEast = fmt['ChineseFont']
    doc.Styles(fmt['Name']).Font.NameAscii = fmt['WesternFont']
    doc.Styles(fmt['Name']).Font.NameOther = fmt['WesternFont']
    doc.Styles(fmt['Name']).Font.Name = fmt['WesternFont']
    doc.Styles(fmt['Name']).Font.Size = fontPoints[fmt['FontSize']]
    doc.Styles(fmt['Name']).Font.Bold = fmt['Bold']
    doc.Styles(fmt['Name']).Font.Italic = fmt['Italic']
    doc.Styles(fmt['Name']).Font.Underline = 0  # wdUnderlineNone
    doc.Styles(fmt['Name']).Font.UnderlineColor = -16777216  # wdColorAutomatic
    doc.Styles(fmt['Name']).Font.StrikeThrough = False
    doc.Styles(fmt['Name']).Font.DoubleStrikeThrough = False
    doc.Styles(fmt['Name']).Font.Outline = False
    doc.Styles(fmt['Name']).Font.Emboss = False
    doc.Styles(fmt['Name']).Font.Shadow = False
    doc.Styles(fmt['Name']).Font.Hidden = False
    doc.Styles(fmt['Name']).Font.SmallCaps = False
    doc.Styles(fmt['Name']).Font.AllCaps = False
    doc.Styles(fmt['Name']).Font.Color = 0  # wdColorBlack
    doc.Styles(fmt['Name']).Font.Engrave = False
    doc.Styles(fmt['Name']).Font.Superscript = False
    doc.Styles(fmt['Name']).Font.Subscript = False
    doc.Styles(fmt['Name']).Font.Scaling = 100
    doc.Styles(fmt['Name']).Font.Kerning = 1
    doc.Styles(fmt['Name']).Font.Animation = 0  # wdAnimationNone
    doc.Styles(fmt['Name']).Font.DisableCharacterSpaceGrid = False
    doc.Styles(fmt['Name']).Font.EmphasisMark = 0  # wdEmphasisMarkNone
    doc.Styles(fmt['Name']).Font.Ligatures = 0  # wdLigaturesNone
    doc.Styles(fmt['Name']).Font.NumberSpacing = 0  # wdNumberSpacingDefault
    doc.Styles(fmt['Name']).Font.NumberForm = 0  # wdNumberFormDefault
    doc.Styles(fmt['Name']).Font.StylisticSet = 0  # wdStylisticSetDefault
    doc.Styles(fmt['Name']).Font.ContextualAlternates = 0
    doc.Styles(fmt['Name']).ParagraphFormat.LeftIndent = 0
    doc.Styles(fmt['Name']).ParagraphFormat.RightIndent = 0
    doc.Styles(fmt['Name']).ParagraphFormat.SpaceBefore = 0
    doc.Styles(fmt['Name']).ParagraphFormat.SpaceBeforeAuto = False
    doc.Styles(fmt['Name']).ParagraphFormat.SpaceAfter = 0
    doc.Styles(fmt['Name']).ParagraphFormat.SpaceAfterAuto = False
    doc.Styles(fmt['Name']).ParagraphFormat.LineSpacingRule = 5  # wdLineSpaceMultiple
    doc.Styles(fmt['Name']).ParagraphFormat.LineSpacing = 12*fmt['LineSpacing']
    if fmt['Alignment'] == '两端':
        doc.Styles(fmt['Name']).ParagraphFormat.Alignment = 3  # wdAlignParagraphJustify
    elif fmt['Alignment'] == '左':
        doc.Styles(fmt['Name']).ParagraphFormat.Alignment = 0
    elif fmt['Alignment'] == '中':
        doc.Styles(fmt['Name']).ParagraphFormat.Alignment = 1
    elif fmt['Alignment'] == '右':
        doc.Styles(fmt['Name']).ParagraphFormat.Alignment = 2
    doc.Styles(fmt['Name']).ParagraphFormat.WidowControl = False
    doc.Styles(fmt['Name']).ParagraphFormat.KeepWithNext = False
    doc.Styles(fmt['Name']).ParagraphFormat.KeepTogether = False
    doc.Styles(fmt['Name']).ParagraphFormat.PageBreakBefore = False
    doc.Styles(fmt['Name']).ParagraphFormat.NoLineNumber = False
    doc.Styles(fmt['Name']).ParagraphFormat.Hyphenation = True
    # doc.Styles(fmt['Name']).ParagraphFormat.FirstLineIndent = fmt['FirstLineIndent']
    doc.Styles(fmt['Name']).ParagraphFormat.OutlineLevel = fmt['OutlineLevel']  # wdOutlineLevelBodyText
    doc.Styles(fmt['Name']).ParagraphFormat.CharacterUnitLeftIndent = fmt['LeftIndent']
    doc.Styles(fmt['Name']).ParagraphFormat.CharacterUnitRightIndent = fmt['RightIndent']
    doc.Styles(fmt['Name']).ParagraphFormat.CharacterUnitFirstLineIndent = fmt['FirstLineIndent']
    doc.Styles(fmt['Name']).ParagraphFormat.LineUnitBefore = fmt['LineUnitBefore']
    doc.Styles(fmt['Name']).ParagraphFormat.LineUnitAfter = fmt['LineUnitAfter']
    doc.Styles(fmt['Name']).ParagraphFormat.MirrorIndents = False
    doc.Styles(fmt['Name']).ParagraphFormat.TextboxTightWrap = 0  # wdTightNone
    doc.Styles(fmt['Name']).ParagraphFormat.AutoAdjustRightIndent = True
    doc.Styles(fmt['Name']).ParagraphFormat.DisableLineHeightGrid = False
    doc.Styles(fmt['Name']).ParagraphFormat.FarEastLineBreakControl = True
    doc.Styles(fmt['Name']).ParagraphFormat.WordWrap = True
    doc.Styles(fmt['Name']).ParagraphFormat.HangingPunctuation = True
    doc.Styles(fmt['Name']).ParagraphFormat.HalfWidthPunctuationOnTopOfLine = False
    doc.Styles(fmt['Name']).ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = True
    doc.Styles(fmt['Name']).ParagraphFormat.AddSpaceBetweenFarEastAndDigit = True
    doc.Styles(fmt['Name']).ParagraphFormat.BaseLineAlignment = 4  # wdBaselineAlignAuto
    doc.Styles(fmt['Name']).NoSpaceBetweenParagraphsOfSameStyle = False
    doc.Styles(fmt['Name']).ParagraphFormat.TabStops.ClearAll
    doc.Styles(fmt['Name']).ParagraphFormat.Shading.Texture = 0  # wdTextureNone
    doc.Styles(fmt['Name']).ParagraphFormat.Shading.ForegroundPatternColor = -16777216  # wdColorAutomatic
    doc.Styles(fmt['Name']).ParagraphFormat.Shading.BackgroundPatternColor = -16777216  # wdColorAutomatic
    doc.Styles(fmt['Name']).ParagraphFormat.Borders(-2).LineStyle = 0  # wdLineStyleNone
    doc.Styles(fmt['Name']).ParagraphFormat.Borders(-4).LineStyle = 0  # wdLineStyleNone
    doc.Styles(fmt['Name']).ParagraphFormat.Borders(-1).LineStyle = 0  # wdLineStyleNone
    doc.Styles(fmt['Name']).ParagraphFormat.Borders(-3).LineStyle = 0  # wdLineStyleNone
    doc.Styles(fmt['Name']).ParagraphFormat.Borders.DistanceFromTop = 1
    doc.Styles(fmt['Name']).ParagraphFormat.Borders.DistanceFromLeft = 4
    doc.Styles(fmt['Name']).ParagraphFormat.Borders.DistanceFromBottom = 1
    doc.Styles(fmt['Name']).ParagraphFormat.Borders.DistanceFromRight = 4
    doc.Styles(fmt['Name']).ParagraphFormat.Borders.Shadow = False
    doc.Styles(fmt['Name']).NoProofing = False
    doc.Styles(fmt['Name']).Frame.Delete
    keys = fmt['Shortcut'].split('+')
    app.CustomizationContext = doc
    if len(keys) == 2:
        # app.FindKey(app.BuildKeyCode(keyCode[keys[0]], keyCode[keys[1]])).Disable
        app.KeyBindings.Add(5, fmt['Name'], app.BuildKeyCode(keyCode[keys[0]], keyCode[keys[1]]))
    elif len(keys) == 3:
        # app.FindKey(app.BuildKeyCode(keyCode[keys[0]], keyCode[keys[1]], keyCode[keys[2]])).Disable
        app.KeyBindings.Add(5, fmt['Name'], app.BuildKeyCode(keyCode[keys[0]], keyCode[keys[1]], keyCode[keys[2]]))
    print('Load', fmt['Name'], 'successfully!')


# word = win32.gencache.EnsureDispatch('Word.Application')
# try:
#     doc = word.ActiveDocument
# except pywintypes.com_error as e:
#     print('com_error:', e)
#     sys.exit(1)
# fmt = {'Name': '正文',
#        'ChineseFont': '宋体',
#        'WesternFont': 'Times New Roman',
#        'FontSize': '小四',
#        'Bold': False,
#        'Italic': False,
#        'Alignment': '两端',
#        'OutlineLevel': 10,
#        'LeftIndent': 0,
#        'RightIndent': 0,
#        'FirstLineIndent': 0,
#        'LineUnitBefore': 0,
#        'LineUnitAfter': 0,
#        'LineSpacing': 1.25,
#        'Shortcut': 'ctrl+`'
#        }
# newStyle(word, doc, fmt)

if __name__ == '__main__':
    loadStylesFromCsv(sys.argv[1])