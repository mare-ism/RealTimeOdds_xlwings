import pandas as pd


# 最高オッズと最低オッズを分離する関数
def get_low_high_odds(odds_range_list):
    odds_low  = [i[0]  for i in odds_range_list]
    odds_high = [i[-1] for i in odds_range_list]
    
    return(odds_low, odds_high)


def get_tansyo_odds(wb):
    # 単勝・複勝シートに接続
    tansyo_sheet = wb.sheets['単勝・複勝']

    # 馬番
    umaban = tansyo_sheet.range('B14:B31').value
    umaban = [0 if i == '' else int(i) for i in umaban]

    # 単勝オッズ
    tansyo_odds = tansyo_sheet.range('D14:D31').value
    tansyo_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in tansyo_odds]

    # データフレーム化
    tansyo_odds_df = pd.DataFrame({
        'umaban': umaban,
        'tansyo_odds': tansyo_odds
    })

    tansyo_odds_df = tansyo_odds_df.dropna()
    tansyo_odds_df.reset_index(drop=True, inplace=True)

    return(tansyo_odds_df)


def get_fukusyo_odds(wb):
    # 単勝・複勝シートに接続
    tansyo_sheet = wb.sheets['単勝・複勝']

    # 馬番
    umaban = tansyo_sheet.range('B14:B31').value
    umaban = [0 if i == '' else int(i) for i in umaban]

    # 複勝オッズ
    fukusyo_odds = tansyo_sheet.range('F14:H31').value
    fukusyo_odds_low, fukusyo_odds_high = get_low_high_odds(fukusyo_odds)
    fukusyo_odds_low  = [1000000.0 if i == '票数なし' else None if i == '' else i for i in fukusyo_odds_low]
    fukusyo_odds_high = [1000000.0 if i == '票数なし' else None if i == '' else i for i in fukusyo_odds_high]

    # データフレーム化
    fukusyo_odds_df = pd.DataFrame({
        'umaban': umaban,
        'fukusyo_odds_low': fukusyo_odds_low,
        'fukusyo_odds_high': fukusyo_odds_high
    })

    fukusyo_odds_df = fukusyo_odds_df.dropna()
    fukusyo_odds_df.reset_index(drop=True, inplace=True)

    return(fukusyo_odds_df)


def get_wakuren_odds(wb):
    # 枠連シート に接続
    wakuren_sheet = wb.sheets['枠連']

    # 枠番
    waku_1 = sum([[i] * (9 - i) for i in range(1, 9)], [])
    waku_2 = sum([list(range(i, 9)) for i in range(1, 9)], [])

    # 枠連オッズ
    wakuren_odds_1 = wakuren_sheet.range('B5:B12').value
    wakuren_odds_2 = wakuren_sheet.range('F5:F11').value
    wakuren_odds_3 = wakuren_sheet.range('J5:J10').value
    wakuren_odds_4 = wakuren_sheet.range('N5:N9').value
    wakuren_odds_5 = wakuren_sheet.range('R5:R8').value
    wakuren_odds_6 = wakuren_sheet.range('V5:V7').value
    wakuren_odds_7 = wakuren_sheet.range('Z5:Z6').value
    wakuren_odds_8 = [wakuren_sheet.range('AD5:AD5').value]

    wakuren_odds = wakuren_odds_1 + wakuren_odds_2 + wakuren_odds_3 + \
                   wakuren_odds_4 + wakuren_odds_5 + wakuren_odds_6 + \
                   wakuren_odds_7 + wakuren_odds_8
    wakuren_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in wakuren_odds]

    # データフレーム化
    wakuren_odds_df = pd.DataFrame({
        'waku_1': waku_1,
        'waku_2': waku_2,
        'wakuren_odds': wakuren_odds
    })

    wakuren_odds_df = wakuren_odds_df.dropna()
    wakuren_odds_df.reset_index(drop=True, inplace=True)

    return(wakuren_odds_df)


def get_umaren_odds(wb):
    # 馬連シート に接続
    umaren_sheet = wb.sheets['馬連']

    # 馬番
    umaban_1 = sum([[i] * (18 - i) for i in range(1, 18)], [])
    umaban_2 = sum([list(range(i, 19)) for i in range(2, 19)], [])

    # 馬連オッズ
    umaren_odds_1  = umaren_sheet.range('B5:B21').value
    umaren_odds_2  = umaren_sheet.range('F5:F20').value
    umaren_odds_3  = umaren_sheet.range('J5:J19').value
    umaren_odds_4  = umaren_sheet.range('N5:N18').value
    umaren_odds_5  = umaren_sheet.range('R5:R17').value
    umaren_odds_6  = umaren_sheet.range('V5:V16').value
    umaren_odds_7  = umaren_sheet.range('Z5:Z15').value
    umaren_odds_8  = umaren_sheet.range('AD5:AD14').value
    umaren_odds_9  = umaren_sheet.range('AH5:AH13').value
    umaren_odds_10 = umaren_sheet.range('B26:B33').value
    umaren_odds_11 = umaren_sheet.range('F26:F32').value
    umaren_odds_12 = umaren_sheet.range('J26:J31').value
    umaren_odds_13 = umaren_sheet.range('N26:N30').value
    umaren_odds_14 = umaren_sheet.range('R26:R29').value
    umaren_odds_15 = umaren_sheet.range('V26:V28').value
    umaren_odds_16 = umaren_sheet.range('Z26:Z27').value
    umaren_odds_17 = [umaren_sheet.range('AD26:AD26').value]

    umaren_odds = umaren_odds_1  + umaren_odds_2  + umaren_odds_3  + \
                  umaren_odds_4  + umaren_odds_5  + umaren_odds_6  + \
                  umaren_odds_7  + umaren_odds_8  + umaren_odds_9  + \
                  umaren_odds_10 + umaren_odds_11 + umaren_odds_12 + \
                  umaren_odds_13 + umaren_odds_14 + umaren_odds_15 + \
                  umaren_odds_16 + umaren_odds_17
    umaren_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in umaren_odds]

    # データフレーム化
    umaren_odds_df = pd.DataFrame({
        'umaban_1': umaban_1,
        'umaban_2': umaban_2,
        'umaren_odds': umaren_odds
    })

    umaren_odds_df = umaren_odds_df.dropna()
    umaren_odds_df.reset_index(drop=True, inplace=True)

    return(umaren_odds_df)


def get_wide_odds(wb):
    try:
        # ワイドシート に接続
        wide_sheet = wb.sheets['ワイド']
    except:
        print('[ワイド]シートが存在しません。\nリアルタイムオッズEX_Pro を使用して下さい。')
        return

    # 馬番
    umaban_1 = sum([[i] * (18 - i) for i in range(1, 18)], [])
    umaban_2 = sum([list(range(i, 19)) for i in range(2, 19)], [])

    # ワイドオッズ
    wide_odds_1  = wide_sheet.range('B5:D21').value
    wide_odds_2  = wide_sheet.range('H5:J20').value
    wide_odds_3  = wide_sheet.range('N5:P19').value
    wide_odds_4  = wide_sheet.range('T5:V18').value
    wide_odds_5  = wide_sheet.range('Z5:AB17').value
    wide_odds_6  = wide_sheet.range('AF5:AH16').value
    wide_odds_7  = wide_sheet.range('B26:D36').value
    wide_odds_8  = wide_sheet.range('H26:J35').value
    wide_odds_9  = wide_sheet.range('N26:P34').value
    wide_odds_10 = wide_sheet.range('T26:V33').value
    wide_odds_11 = wide_sheet.range('Z26:AB32').value
    wide_odds_12 = wide_sheet.range('AF26:AH31').value
    wide_odds_13 = wide_sheet.range('B41:D45').value
    wide_odds_14 = wide_sheet.range('H41:J44').value
    wide_odds_15 = wide_sheet.range('N41:P43').value
    wide_odds_16 = wide_sheet.range('T41:V42').value
    wide_odds_17 = wide_sheet.range('Z41:AB41').value

    wide_odds_low_1,  wide_odds_high_1  = get_low_high_odds(wide_odds_1)
    wide_odds_low_2,  wide_odds_high_2  = get_low_high_odds(wide_odds_2)
    wide_odds_low_3,  wide_odds_high_3  = get_low_high_odds(wide_odds_3)
    wide_odds_low_4,  wide_odds_high_4  = get_low_high_odds(wide_odds_4)
    wide_odds_low_5,  wide_odds_high_5  = get_low_high_odds(wide_odds_5)
    wide_odds_low_6,  wide_odds_high_6  = get_low_high_odds(wide_odds_6)
    wide_odds_low_7,  wide_odds_high_7  = get_low_high_odds(wide_odds_7)
    wide_odds_low_8,  wide_odds_high_8  = get_low_high_odds(wide_odds_8)
    wide_odds_low_9,  wide_odds_high_9  = get_low_high_odds(wide_odds_9)
    wide_odds_low_10, wide_odds_high_10 = get_low_high_odds(wide_odds_10)
    wide_odds_low_11, wide_odds_high_11 = get_low_high_odds(wide_odds_11)
    wide_odds_low_12, wide_odds_high_12 = get_low_high_odds(wide_odds_12)
    wide_odds_low_13, wide_odds_high_13 = get_low_high_odds(wide_odds_13)
    wide_odds_low_14, wide_odds_high_14 = get_low_high_odds(wide_odds_14)
    wide_odds_low_15, wide_odds_high_15 = get_low_high_odds(wide_odds_15)
    wide_odds_low_16, wide_odds_high_16 = get_low_high_odds(wide_odds_16)
    wide_odds_low_17, wide_odds_high_17 = [wide_odds_17[0]], [wide_odds_17[-1]]

    wide_odds_low = wide_odds_low_1  + wide_odds_low_2  + wide_odds_low_3  + \
                    wide_odds_low_4  + wide_odds_low_5  + wide_odds_low_6  + \
                    wide_odds_low_7  + wide_odds_low_8  + wide_odds_low_9  + \
                    wide_odds_low_10 + wide_odds_low_11 + wide_odds_low_12 + \
                    wide_odds_low_13 + wide_odds_low_14 + wide_odds_low_15 + \
                    wide_odds_low_16 + wide_odds_low_17
    wide_odds_low = [1000000.0 if i == '票数なし' else None if i == '' else i for i in wide_odds_low]

    wide_odds_high = wide_odds_high_1  + wide_odds_high_2  + wide_odds_high_3  + \
                     wide_odds_high_4  + wide_odds_high_5  + wide_odds_high_6  + \
                     wide_odds_high_7  + wide_odds_high_8  + wide_odds_high_9  + \
                     wide_odds_high_10 + wide_odds_high_11 + wide_odds_high_12 + \
                     wide_odds_high_13 + wide_odds_high_14 + wide_odds_high_15 + \
                     wide_odds_high_16 + wide_odds_high_17
    wide_odds_high = [1000000.0 if i == '票数なし' else None if i == '' else i for i in wide_odds_high]

    # データフレーム化
    wide_odds_df = pd.DataFrame({
        'umaban_1': umaban_1,
        'umaban_2': umaban_2,
        'wide_odds_low': wide_odds_low,
        'wide_odds_high': wide_odds_high
    })

    wide_odds_df = wide_odds_df.dropna()
    wide_odds_df.reset_index(drop=True, inplace=True)

    return(wide_odds_df)


def get_umatan_odds(wb):
    try:
        # 馬単シート に接続
        umatan_sheet = wb.sheets['馬単']
    except:
        print('[馬単]シートが存在しません。\nリアルタイムオッズEX_Pro を使用して下さい。')
        return

    # 馬番
    umaban_1 = sum([[i] * 18 for i in range(1, 19)], [])
    umaban_2 = sum([list(range(1, 19)) for i in range(1, 19)], [])

    # 馬単オッズ
    umatan_odds_1  = umatan_sheet.range('B5:B22').value
    umatan_odds_2  = umatan_sheet.range('F5:F22').value
    umatan_odds_3  = umatan_sheet.range('J5:J22').value
    umatan_odds_4  = umatan_sheet.range('N5:N22').value
    umatan_odds_5  = umatan_sheet.range('R5:R22').value
    umatan_odds_6  = umatan_sheet.range('V5:V22').value
    umatan_odds_7  = umatan_sheet.range('Z5:Z22').value
    umatan_odds_8  = umatan_sheet.range('AD5:AD22').value
    umatan_odds_9  = umatan_sheet.range('AH5:AH22').value
    umatan_odds_10 = umatan_sheet.range('B27:B44').value
    umatan_odds_11 = umatan_sheet.range('F27:F44').value
    umatan_odds_12 = umatan_sheet.range('J27:J44').value
    umatan_odds_13 = umatan_sheet.range('N27:N44').value
    umatan_odds_14 = umatan_sheet.range('R27:R44').value
    umatan_odds_15 = umatan_sheet.range('V27:V44').value
    umatan_odds_16 = umatan_sheet.range('Z27:Z44').value
    umatan_odds_17 = umatan_sheet.range('AD27:AD44').value
    umatan_odds_18 = umatan_sheet.range('AH27:AH44').value

    umatan_odds = umatan_odds_1  + umatan_odds_2  + umatan_odds_3  + \
                  umatan_odds_4  + umatan_odds_5  + umatan_odds_6  + \
                  umatan_odds_7  + umatan_odds_8  + umatan_odds_9  + \
                  umatan_odds_10 + umatan_odds_11 + umatan_odds_12 + \
                  umatan_odds_13 + umatan_odds_14 + umatan_odds_15 + \
                  umatan_odds_16 + umatan_odds_17 + umatan_odds_18
    umatan_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in umatan_odds]


    # データフレーム化
    umatan_odds_df = pd.DataFrame({
        'umaban_1': umaban_1,
        'umaban_2': umaban_2,
        'umatan_odds': umatan_odds
    })

    umatan_odds_df = umatan_odds_df.query('umaban_1 != umaban_2').dropna()
    umatan_odds_df.reset_index(drop=True, inplace=True)

    return(umatan_odds_df)


def get_sanrenpuku_odds(wb):
    try:
        # 3連複シート に接続
        sanrenpuku_sheet = wb.sheets['3連複']
    except:
        print('[3連複]シートが存在しません。\nリアルタイムオッズEX_Pro を使用して下さい。')
        return

    # 馬番
    umaban_1 = sum([[i] * int((18 - i - 1) * (18 - i) / 2) for i in range(1, 18)], [])
    umaban_2 = sum([[i + j] * (18 - i - j) for j in range(16) for i in range(2, 18 - j)], [])
    umaban_3 = sum([list(range(i, 19)) for j in range(16) for i in range(3 + j, 19)], [])

    # 3連複オッズ
    sanrenpuku_odds_1_2   =  sanrenpuku_sheet.range('B7:B22').value
    sanrenpuku_odds_1_3   =  sanrenpuku_sheet.range('F7:F21').value
    sanrenpuku_odds_1_4   =  sanrenpuku_sheet.range('J7:J20').value
    sanrenpuku_odds_1_5   =  sanrenpuku_sheet.range('N7:N19').value
    sanrenpuku_odds_1_6   =  sanrenpuku_sheet.range('R7:R18').value
    sanrenpuku_odds_1_7   =  sanrenpuku_sheet.range('V7:V17').value
    sanrenpuku_odds_1_8   =  sanrenpuku_sheet.range('Z7:Z16').value
    sanrenpuku_odds_1_9   =  sanrenpuku_sheet.range('AD7:AD15').value
    sanrenpuku_odds_1_10  =  sanrenpuku_sheet.range('AH7:AH14').value
    sanrenpuku_odds_1_11  =  sanrenpuku_sheet.range('AL7:AL13').value
    sanrenpuku_odds_1_12  =  sanrenpuku_sheet.range('AP7:AP12').value
    sanrenpuku_odds_1_13  =  sanrenpuku_sheet.range('AT7:AT11').value
    sanrenpuku_odds_1_14  =  sanrenpuku_sheet.range('AX7:AX10').value
    sanrenpuku_odds_1_15  =  sanrenpuku_sheet.range('BB7:BB9').value
    sanrenpuku_odds_1_16  =  sanrenpuku_sheet.range('BF7:BF8').value
    sanrenpuku_odds_1_17  = [sanrenpuku_sheet.range('BJ7:BJ7').value]
    sanrenpuku_odds_2_3   =  sanrenpuku_sheet.range('B28:B42').value
    sanrenpuku_odds_2_4   =  sanrenpuku_sheet.range('F28:F41').value
    sanrenpuku_odds_2_5   =  sanrenpuku_sheet.range('J28:J40').value
    sanrenpuku_odds_2_6   =  sanrenpuku_sheet.range('N28:N39').value
    sanrenpuku_odds_2_7   =  sanrenpuku_sheet.range('R28:R38').value
    sanrenpuku_odds_2_8   =  sanrenpuku_sheet.range('V28:V37').value
    sanrenpuku_odds_2_9   =  sanrenpuku_sheet.range('Z28:Z36').value
    sanrenpuku_odds_2_10  =  sanrenpuku_sheet.range('AD28:AD35').value
    sanrenpuku_odds_2_11  =  sanrenpuku_sheet.range('AH28:AH34').value
    sanrenpuku_odds_2_12  =  sanrenpuku_sheet.range('AL28:AL33').value
    sanrenpuku_odds_2_13  =  sanrenpuku_sheet.range('AP28:AP32').value
    sanrenpuku_odds_2_14  =  sanrenpuku_sheet.range('AT28:AT31').value
    sanrenpuku_odds_2_15  =  sanrenpuku_sheet.range('AX28:AX30').value
    sanrenpuku_odds_2_16  =  sanrenpuku_sheet.range('BB28:BB29').value
    sanrenpuku_odds_2_17  = [sanrenpuku_sheet.range('BF28:BF28').value]
    sanrenpuku_odds_3_4   =  sanrenpuku_sheet.range('B48:B61').value
    sanrenpuku_odds_3_5   =  sanrenpuku_sheet.range('F48:F60').value
    sanrenpuku_odds_3_6   =  sanrenpuku_sheet.range('J48:J59').value
    sanrenpuku_odds_3_7   =  sanrenpuku_sheet.range('N48:N58').value
    sanrenpuku_odds_3_8   =  sanrenpuku_sheet.range('R48:R57').value
    sanrenpuku_odds_3_9   =  sanrenpuku_sheet.range('V48:V56').value
    sanrenpuku_odds_3_10  =  sanrenpuku_sheet.range('Z48:Z55').value
    sanrenpuku_odds_3_11  =  sanrenpuku_sheet.range('AD48:AD54').value
    sanrenpuku_odds_3_12  =  sanrenpuku_sheet.range('AH48:AH53').value
    sanrenpuku_odds_3_13  =  sanrenpuku_sheet.range('AL48:AL52').value
    sanrenpuku_odds_3_14  =  sanrenpuku_sheet.range('AP48:AP51').value
    sanrenpuku_odds_3_15  =  sanrenpuku_sheet.range('AT48:AT50').value
    sanrenpuku_odds_3_16  =  sanrenpuku_sheet.range('AX48:AX49').value
    sanrenpuku_odds_3_17  = [sanrenpuku_sheet.range('BB48:BB48').value]
    sanrenpuku_odds_4_5   =  sanrenpuku_sheet.range('B67:B79').value
    sanrenpuku_odds_4_6   =  sanrenpuku_sheet.range('F67:F78').value
    sanrenpuku_odds_4_7   =  sanrenpuku_sheet.range('J67:J77').value
    sanrenpuku_odds_4_8   =  sanrenpuku_sheet.range('N67:N76').value
    sanrenpuku_odds_4_9   =  sanrenpuku_sheet.range('R67:R75').value
    sanrenpuku_odds_4_10  =  sanrenpuku_sheet.range('V67:V74').value
    sanrenpuku_odds_4_11  =  sanrenpuku_sheet.range('Z67:Z73').value
    sanrenpuku_odds_4_12  =  sanrenpuku_sheet.range('AD67:AD72').value
    sanrenpuku_odds_4_13  =  sanrenpuku_sheet.range('AH67:AH71').value
    sanrenpuku_odds_4_14  =  sanrenpuku_sheet.range('AL67:AL70').value
    sanrenpuku_odds_4_15  =  sanrenpuku_sheet.range('AP67:AP69').value
    sanrenpuku_odds_4_16  =  sanrenpuku_sheet.range('AT67:AT68').value
    sanrenpuku_odds_4_17  = [sanrenpuku_sheet.range('AX67:AX67').value]
    sanrenpuku_odds_5_6   =  sanrenpuku_sheet.range('B85:B96').value
    sanrenpuku_odds_5_7   =  sanrenpuku_sheet.range('F85:F95').value
    sanrenpuku_odds_5_8   =  sanrenpuku_sheet.range('J85:J94').value
    sanrenpuku_odds_5_9   =  sanrenpuku_sheet.range('N85:N93').value
    sanrenpuku_odds_5_10  =  sanrenpuku_sheet.range('R85:R92').value
    sanrenpuku_odds_5_11  =  sanrenpuku_sheet.range('V85:V91').value
    sanrenpuku_odds_5_12  =  sanrenpuku_sheet.range('Z85:Z90').value
    sanrenpuku_odds_5_13  =  sanrenpuku_sheet.range('AD85:AD89').value
    sanrenpuku_odds_5_14  =  sanrenpuku_sheet.range('AH85:AH88').value
    sanrenpuku_odds_5_15  =  sanrenpuku_sheet.range('AL85:AL87').value
    sanrenpuku_odds_5_16  =  sanrenpuku_sheet.range('AP85:AP86').value
    sanrenpuku_odds_5_17  = [sanrenpuku_sheet.range('AT85:AT85').value]
    sanrenpuku_odds_6_7   =  sanrenpuku_sheet.range('B102:B112').value
    sanrenpuku_odds_6_8   =  sanrenpuku_sheet.range('F102:F111').value
    sanrenpuku_odds_6_9   =  sanrenpuku_sheet.range('J102:J110').value
    sanrenpuku_odds_6_10  =  sanrenpuku_sheet.range('N102:N109').value
    sanrenpuku_odds_6_11  =  sanrenpuku_sheet.range('R102:R108').value
    sanrenpuku_odds_6_12  =  sanrenpuku_sheet.range('V102:V107').value
    sanrenpuku_odds_6_13  =  sanrenpuku_sheet.range('Z102:Z106').value
    sanrenpuku_odds_6_14  =  sanrenpuku_sheet.range('AD102:AD105').value
    sanrenpuku_odds_6_15  =  sanrenpuku_sheet.range('AH102:AH104').value
    sanrenpuku_odds_6_16  =  sanrenpuku_sheet.range('AL102:AL103').value
    sanrenpuku_odds_6_17  = [sanrenpuku_sheet.range('AP102:AP102').value]
    sanrenpuku_odds_7_8   =  sanrenpuku_sheet.range('B118:B127').value
    sanrenpuku_odds_7_9   =  sanrenpuku_sheet.range('F118:F126').value
    sanrenpuku_odds_7_10  =  sanrenpuku_sheet.range('J118:J125').value
    sanrenpuku_odds_7_11  =  sanrenpuku_sheet.range('N118:N124').value
    sanrenpuku_odds_7_12  =  sanrenpuku_sheet.range('R118:R123').value
    sanrenpuku_odds_7_13  =  sanrenpuku_sheet.range('V118:V122').value
    sanrenpuku_odds_7_14  =  sanrenpuku_sheet.range('Z118:Z121').value
    sanrenpuku_odds_7_15  =  sanrenpuku_sheet.range('AD118:AD120').value
    sanrenpuku_odds_7_16  =  sanrenpuku_sheet.range('AH118:AH119').value
    sanrenpuku_odds_7_17  = [sanrenpuku_sheet.range('AL118:AL118').value]
    sanrenpuku_odds_8_9   =  sanrenpuku_sheet.range('B133:B141').value
    sanrenpuku_odds_8_10  =  sanrenpuku_sheet.range('F133:F140').value
    sanrenpuku_odds_8_11  =  sanrenpuku_sheet.range('J133:J139').value
    sanrenpuku_odds_8_12  =  sanrenpuku_sheet.range('N133:N138').value
    sanrenpuku_odds_8_13  =  sanrenpuku_sheet.range('R133:R137').value
    sanrenpuku_odds_8_14  =  sanrenpuku_sheet.range('V133:V136').value
    sanrenpuku_odds_8_15  =  sanrenpuku_sheet.range('Z133:Z135').value
    sanrenpuku_odds_8_16  =  sanrenpuku_sheet.range('AD133:AD134').value
    sanrenpuku_odds_8_17  = [sanrenpuku_sheet.range('AH133:AH133').value]
    sanrenpuku_odds_9_10  =  sanrenpuku_sheet.range('B147:B154').value
    sanrenpuku_odds_9_11  =  sanrenpuku_sheet.range('F147:F153').value
    sanrenpuku_odds_9_12  =  sanrenpuku_sheet.range('J147:J152').value
    sanrenpuku_odds_9_13  =  sanrenpuku_sheet.range('N147:N151').value
    sanrenpuku_odds_9_14  =  sanrenpuku_sheet.range('R147:R150').value
    sanrenpuku_odds_9_15  =  sanrenpuku_sheet.range('V147:V149').value
    sanrenpuku_odds_9_16  =  sanrenpuku_sheet.range('Z147:Z148').value
    sanrenpuku_odds_9_17  = [sanrenpuku_sheet.range('AD147:AD147').value]
    sanrenpuku_odds_10_11 =  sanrenpuku_sheet.range('B160:B166').value
    sanrenpuku_odds_10_12 =  sanrenpuku_sheet.range('F160:F165').value
    sanrenpuku_odds_10_13 =  sanrenpuku_sheet.range('J160:J164').value
    sanrenpuku_odds_10_14 =  sanrenpuku_sheet.range('N160:N163').value
    sanrenpuku_odds_10_15 =  sanrenpuku_sheet.range('R160:R162').value
    sanrenpuku_odds_10_16 =  sanrenpuku_sheet.range('V160:V161').value
    sanrenpuku_odds_10_17 = [sanrenpuku_sheet.range('Z160:Z160').value]
    sanrenpuku_odds_11_12 =  sanrenpuku_sheet.range('B172:B177').value
    sanrenpuku_odds_11_13 =  sanrenpuku_sheet.range('F172:F176').value
    sanrenpuku_odds_11_14 =  sanrenpuku_sheet.range('J172:J175').value
    sanrenpuku_odds_11_15 =  sanrenpuku_sheet.range('N172:N174').value
    sanrenpuku_odds_11_16 =  sanrenpuku_sheet.range('R172:R173').value
    sanrenpuku_odds_11_17 = [sanrenpuku_sheet.range('V172:V172').value]
    sanrenpuku_odds_12_13 =  sanrenpuku_sheet.range('B183:B187').value
    sanrenpuku_odds_12_14 =  sanrenpuku_sheet.range('F183:F186').value
    sanrenpuku_odds_12_15 =  sanrenpuku_sheet.range('J183:J185').value
    sanrenpuku_odds_12_16 =  sanrenpuku_sheet.range('N183:N184').value
    sanrenpuku_odds_12_17 = [sanrenpuku_sheet.range('R183:R183').value]
    sanrenpuku_odds_13_14 =  sanrenpuku_sheet.range('B193:B196').value
    sanrenpuku_odds_13_15 =  sanrenpuku_sheet.range('F193:F195').value
    sanrenpuku_odds_13_16 =  sanrenpuku_sheet.range('J193:J194').value
    sanrenpuku_odds_13_17 = [sanrenpuku_sheet.range('N193:N193').value]
    sanrenpuku_odds_14_15 =  sanrenpuku_sheet.range('B202:B204').value
    sanrenpuku_odds_14_16 =  sanrenpuku_sheet.range('F202:F203').value
    sanrenpuku_odds_14_17 = [sanrenpuku_sheet.range('J202:J202').value]
    sanrenpuku_odds_15_16 =  sanrenpuku_sheet.range('B210:B211').value
    sanrenpuku_odds_15_17 = [sanrenpuku_sheet.range('F210:F210').value]
    sanrenpuku_odds_16_17 = [sanrenpuku_sheet.range('B217:B217').value]

    sanrenpuku_odds = sanrenpuku_odds_1_2   + sanrenpuku_odds_1_3   + sanrenpuku_odds_1_4   + \
                      sanrenpuku_odds_1_5   + sanrenpuku_odds_1_6   + sanrenpuku_odds_1_7   + \
                      sanrenpuku_odds_1_8   + sanrenpuku_odds_1_9   + sanrenpuku_odds_1_10  + \
                      sanrenpuku_odds_1_11  + sanrenpuku_odds_1_12  + sanrenpuku_odds_1_13  + \
                      sanrenpuku_odds_1_14  + sanrenpuku_odds_1_15  + sanrenpuku_odds_1_16  + \
                      sanrenpuku_odds_1_17  + \
                      sanrenpuku_odds_2_3   + sanrenpuku_odds_2_4   + sanrenpuku_odds_2_5   + \
                      sanrenpuku_odds_2_6   + sanrenpuku_odds_2_7   + sanrenpuku_odds_2_8   + \
                      sanrenpuku_odds_2_9   + sanrenpuku_odds_2_10  + sanrenpuku_odds_2_11  + \
                      sanrenpuku_odds_2_12  + sanrenpuku_odds_2_13  + sanrenpuku_odds_2_14  + \
                      sanrenpuku_odds_2_15  + sanrenpuku_odds_2_16  + sanrenpuku_odds_2_17  + \
                      sanrenpuku_odds_3_4   + sanrenpuku_odds_3_5   + sanrenpuku_odds_3_6   + \
                      sanrenpuku_odds_3_7   + sanrenpuku_odds_3_8   + sanrenpuku_odds_3_9   + \
                      sanrenpuku_odds_3_10  + sanrenpuku_odds_3_11  + sanrenpuku_odds_3_12  + \
                      sanrenpuku_odds_3_13  + sanrenpuku_odds_3_14  + sanrenpuku_odds_3_15  + \
                      sanrenpuku_odds_3_16  + sanrenpuku_odds_3_17  + \
                      sanrenpuku_odds_4_5   + sanrenpuku_odds_4_6   + sanrenpuku_odds_4_7   + \
                      sanrenpuku_odds_4_8   + sanrenpuku_odds_4_9   + sanrenpuku_odds_4_10  + \
                      sanrenpuku_odds_4_11  + sanrenpuku_odds_4_12  + sanrenpuku_odds_4_13  + \
                      sanrenpuku_odds_4_14  + sanrenpuku_odds_4_15  + sanrenpuku_odds_4_16  + \
                      sanrenpuku_odds_4_17  + \
                      sanrenpuku_odds_5_6   + sanrenpuku_odds_5_7   + sanrenpuku_odds_5_8   + \
                      sanrenpuku_odds_5_9   + sanrenpuku_odds_5_10  + sanrenpuku_odds_5_11  + \
                      sanrenpuku_odds_5_12  + sanrenpuku_odds_5_13  + sanrenpuku_odds_5_14  + \
                      sanrenpuku_odds_5_15  + sanrenpuku_odds_5_16  + sanrenpuku_odds_5_17  + \
                      sanrenpuku_odds_6_7   + sanrenpuku_odds_6_8   + sanrenpuku_odds_6_9   + \
                      sanrenpuku_odds_6_10  + sanrenpuku_odds_6_11  + sanrenpuku_odds_6_12  + \
                      sanrenpuku_odds_6_13  + sanrenpuku_odds_6_14  + sanrenpuku_odds_6_15  + \
                      sanrenpuku_odds_6_16  + sanrenpuku_odds_6_17  + \
                      sanrenpuku_odds_7_8   + sanrenpuku_odds_7_9   + sanrenpuku_odds_7_10  + \
                      sanrenpuku_odds_7_11  + sanrenpuku_odds_7_12  + sanrenpuku_odds_7_13  + \
                      sanrenpuku_odds_7_14  + sanrenpuku_odds_7_15  + sanrenpuku_odds_7_16  + \
                      sanrenpuku_odds_7_17  + \
                      sanrenpuku_odds_8_9   + sanrenpuku_odds_8_10  + sanrenpuku_odds_8_11  + \
                      sanrenpuku_odds_8_12  + sanrenpuku_odds_8_13  + sanrenpuku_odds_8_14  + \
                      sanrenpuku_odds_8_15  + sanrenpuku_odds_8_16  + sanrenpuku_odds_8_17  + \
                      sanrenpuku_odds_9_10  + sanrenpuku_odds_9_11  + sanrenpuku_odds_9_12  + \
                      sanrenpuku_odds_9_13  + sanrenpuku_odds_9_14  + sanrenpuku_odds_9_15  + \
                      sanrenpuku_odds_9_16  + sanrenpuku_odds_9_17  + \
                      sanrenpuku_odds_10_11 + sanrenpuku_odds_10_12 + sanrenpuku_odds_10_13 + \
                      sanrenpuku_odds_10_14 + sanrenpuku_odds_10_15 + sanrenpuku_odds_10_16 + \
                      sanrenpuku_odds_10_17 + \
                      sanrenpuku_odds_11_12 + sanrenpuku_odds_11_13 + sanrenpuku_odds_11_14 + \
                      sanrenpuku_odds_11_15 + sanrenpuku_odds_11_16 + sanrenpuku_odds_11_17 + \
                      sanrenpuku_odds_12_13 + sanrenpuku_odds_12_14 + sanrenpuku_odds_12_15 + \
                      sanrenpuku_odds_12_16 + sanrenpuku_odds_12_17 + \
                      sanrenpuku_odds_13_14 + sanrenpuku_odds_13_15 + sanrenpuku_odds_13_16 + \
                      sanrenpuku_odds_13_17 + \
                      sanrenpuku_odds_14_15 + sanrenpuku_odds_14_16 + sanrenpuku_odds_14_17 + \
                      sanrenpuku_odds_15_16 + sanrenpuku_odds_15_17 + \
                      sanrenpuku_odds_16_17
    sanrenpuku_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in sanrenpuku_odds]

    # データフレーム化
    sanrenpuku_odds_df = pd.DataFrame({
        'umaban_1': umaban_1,
        'umaban_2': umaban_2,
        'umaban_3': umaban_3,
        'sanrenpuku_odds': sanrenpuku_odds
    })

    sanrenpuku_odds_df = sanrenpuku_odds_df.dropna()
    sanrenpuku_odds_df.reset_index(drop=True, inplace=True)

    return(sanrenpuku_odds_df)


def get_sanrentan_odds(wb):
    try:
        # 3連単シート に接続
        sanrentan_sheet = wb.sheets['3連単']
    except:
        print('[3連単]シートが存在しません。\nリアルタイムオッズEX_Pro を使用して下さい。')
        return

    # 馬番
    umaban_1 = sum([[i] * 18 * 17 for i in range(1, 19)], [])
    umaban_2 = sum([[i] * 18 for j in range(1, 19) for i in range(1, 19) if i != j], [])
    umaban_3 = [i for i in range(1, 19)] * 17 * 18

    # 3連複オッズ
    offset_rows = 9
    sanrentan_odds_1_2   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_1_3   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_1_4   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_1_5   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_1_6   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_1_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_1_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_1_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_1_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_1_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_1_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_1_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_1_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_1_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_1_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_1_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_1_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_2_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_2_3   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_2_4   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_2_5   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_2_6   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_2_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_2_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_2_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_2_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_2_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_2_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_2_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_2_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_2_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_2_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_2_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_2_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_3_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_3_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_3_4   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_3_5   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_3_6   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_3_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_3_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_3_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_3_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_3_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_3_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_3_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_3_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_3_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_3_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_3_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_3_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_4_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_4_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_4_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_4_5   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_4_6   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_4_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_4_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_4_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_4_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_4_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_4_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_4_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_4_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_4_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_4_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_4_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_4_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_5_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_5_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_5_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_5_4   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_5_6   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_5_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_5_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_5_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_5_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_5_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_5_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_5_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_5_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_5_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_5_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_5_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_5_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_6_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_6_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_6_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_6_4   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_6_5   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_6_7   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_6_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_6_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_6_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_6_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_6_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_6_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_6_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_6_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_6_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_6_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_6_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_7_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_7_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_7_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_7_4   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_7_5   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_7_6   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_7_8   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_7_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_7_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_7_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_7_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_7_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_7_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_7_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_7_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_7_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_7_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_8_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_8_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_8_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_8_4   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_8_5   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_8_6   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_8_7   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_8_9   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_8_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_8_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_8_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_8_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_8_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_8_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_8_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_8_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_8_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_9_1   = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_9_2   = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_9_3   = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_9_4   = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_9_5   = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_9_6   = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_9_7   = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_9_8   = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_9_10  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_9_11  = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_9_12  = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_9_13  = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_9_14  = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_9_15  = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_9_16  = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_9_17  = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_9_18  = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_10_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_10_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_10_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_10_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_10_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_10_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_10_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_10_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_10_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_10_11 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_10_12 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_10_13 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_10_14 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_10_15 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_10_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_10_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_10_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_11_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_11_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_11_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_11_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_11_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_11_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_11_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_11_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_11_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_11_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_11_12 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_11_13 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_11_14 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_11_15 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_11_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_11_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_11_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_12_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_12_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_12_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_12_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_12_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_12_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_12_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_12_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_12_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_12_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_12_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_12_13 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_12_14 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_12_15 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_12_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_12_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_12_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_13_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_13_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_13_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_13_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_13_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_13_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_13_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_13_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_13_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_13_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_13_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_13_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_13_14 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_13_15 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_13_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_13_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_13_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_14_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_14_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_14_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_14_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_14_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_14_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_14_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_14_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_14_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_14_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_14_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_14_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_14_13 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_14_15 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_14_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_14_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_14_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_15_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_15_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_15_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_15_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_15_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_15_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_15_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_15_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_15_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_15_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_15_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_15_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_15_13 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_15_14 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_15_16 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_15_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_15_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_16_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_16_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_16_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_16_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_16_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_16_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_16_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_16_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_16_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_16_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_16_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_16_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_16_13 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_16_14 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_16_15 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_16_17 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_16_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_17_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_17_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_17_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_17_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_17_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_17_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_17_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_17_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_17_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_17_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_17_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_17_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_17_13 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_17_14 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_17_15 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_17_16 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_17_18 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    offset_rows += 25
    sanrentan_odds_18_1  = sanrentan_sheet.range(f'B{offset_rows}:B{offset_rows+17}').value
    sanrentan_odds_18_2  = sanrentan_sheet.range(f'F{offset_rows}:F{offset_rows+17}').value
    sanrentan_odds_18_3  = sanrentan_sheet.range(f'J{offset_rows}:J{offset_rows+17}').value
    sanrentan_odds_18_4  = sanrentan_sheet.range(f'N{offset_rows}:N{offset_rows+17}').value
    sanrentan_odds_18_5  = sanrentan_sheet.range(f'R{offset_rows}:R{offset_rows+17}').value
    sanrentan_odds_18_6  = sanrentan_sheet.range(f'V{offset_rows}:V{offset_rows+17}').value
    sanrentan_odds_18_7  = sanrentan_sheet.range(f'Z{offset_rows}:Z{offset_rows+17}').value
    sanrentan_odds_18_8  = sanrentan_sheet.range(f'AD{offset_rows}:AD{offset_rows+17}').value
    sanrentan_odds_18_9  = sanrentan_sheet.range(f'AH{offset_rows}:AH{offset_rows+17}').value
    sanrentan_odds_18_10 = sanrentan_sheet.range(f'AL{offset_rows}:AL{offset_rows+17}').value
    sanrentan_odds_18_11 = sanrentan_sheet.range(f'AP{offset_rows}:AP{offset_rows+17}').value
    sanrentan_odds_18_12 = sanrentan_sheet.range(f'AT{offset_rows}:AT{offset_rows+17}').value
    sanrentan_odds_18_13 = sanrentan_sheet.range(f'AX{offset_rows}:AX{offset_rows+17}').value
    sanrentan_odds_18_14 = sanrentan_sheet.range(f'BB{offset_rows}:BB{offset_rows+17}').value
    sanrentan_odds_18_15 = sanrentan_sheet.range(f'BF{offset_rows}:BF{offset_rows+17}').value
    sanrentan_odds_18_16 = sanrentan_sheet.range(f'BJ{offset_rows}:BJ{offset_rows+17}').value
    sanrentan_odds_18_17 = sanrentan_sheet.range(f'BN{offset_rows}:BN{offset_rows+17}').value

    sanrentan_odds = sanrentan_odds_1_2   + sanrentan_odds_1_3   + sanrentan_odds_1_4  + \
                     sanrentan_odds_1_5   + sanrentan_odds_1_6   + sanrentan_odds_1_7  + \
                     sanrentan_odds_1_8   + sanrentan_odds_1_9   + sanrentan_odds_1_10 + \
                     sanrentan_odds_1_11  + sanrentan_odds_1_12  + sanrentan_odds_1_13 + \
                     sanrentan_odds_1_14  + sanrentan_odds_1_15  + sanrentan_odds_1_16 + \
                     sanrentan_odds_1_17  + sanrentan_odds_1_18  + \
                     sanrentan_odds_2_1   + sanrentan_odds_2_3   + sanrentan_odds_2_4  + \
                     sanrentan_odds_2_5   + sanrentan_odds_2_6   + sanrentan_odds_2_7  + \
                     sanrentan_odds_2_8   + sanrentan_odds_2_9   + sanrentan_odds_2_10 + \
                     sanrentan_odds_2_11  + sanrentan_odds_2_12  + sanrentan_odds_2_13 + \
                     sanrentan_odds_2_14  + sanrentan_odds_2_15  + sanrentan_odds_2_16 + \
                     sanrentan_odds_2_17  + sanrentan_odds_2_18  + \
                     sanrentan_odds_3_1   + sanrentan_odds_3_2   + sanrentan_odds_3_4  + \
                     sanrentan_odds_3_5   + sanrentan_odds_3_6   + sanrentan_odds_3_7  + \
                     sanrentan_odds_3_8   + sanrentan_odds_3_9   + sanrentan_odds_3_10 + \
                     sanrentan_odds_3_11  + sanrentan_odds_3_12  + sanrentan_odds_3_13 + \
                     sanrentan_odds_3_14  + sanrentan_odds_3_15  + sanrentan_odds_3_16 + \
                     sanrentan_odds_3_17  + sanrentan_odds_3_18  + \
                     sanrentan_odds_4_1   + sanrentan_odds_4_2   + sanrentan_odds_4_3  + \
                     sanrentan_odds_4_5   + sanrentan_odds_4_6   + sanrentan_odds_4_7  + \
                     sanrentan_odds_4_8   + sanrentan_odds_4_9   + sanrentan_odds_4_10 + \
                     sanrentan_odds_4_11  + sanrentan_odds_4_12  + sanrentan_odds_4_13 + \
                     sanrentan_odds_4_14  + sanrentan_odds_4_15  + sanrentan_odds_4_16 + \
                     sanrentan_odds_4_17  + sanrentan_odds_4_18  + \
                     sanrentan_odds_5_1   + sanrentan_odds_5_2   + sanrentan_odds_5_3  + \
                     sanrentan_odds_5_4   + sanrentan_odds_5_6   + sanrentan_odds_5_7  + \
                     sanrentan_odds_5_8   + sanrentan_odds_5_9   + sanrentan_odds_5_10 + \
                     sanrentan_odds_5_11  + sanrentan_odds_5_12  + sanrentan_odds_5_13 + \
                     sanrentan_odds_5_14  + sanrentan_odds_5_15  + sanrentan_odds_5_16 + \
                     sanrentan_odds_5_17  + sanrentan_odds_5_18  + \
                     sanrentan_odds_6_1   + sanrentan_odds_6_2   + sanrentan_odds_6_3  + \
                     sanrentan_odds_6_4   + sanrentan_odds_6_5   + sanrentan_odds_6_7  + \
                     sanrentan_odds_6_8   + sanrentan_odds_6_9   + sanrentan_odds_6_10 + \
                     sanrentan_odds_6_11  + sanrentan_odds_6_12  + sanrentan_odds_6_13 + \
                     sanrentan_odds_6_14  + sanrentan_odds_6_15  + sanrentan_odds_6_16 + \
                     sanrentan_odds_6_17  + sanrentan_odds_6_18  + \
                     sanrentan_odds_7_1   + sanrentan_odds_7_2   + sanrentan_odds_7_3  + \
                     sanrentan_odds_7_4   + sanrentan_odds_7_5   + sanrentan_odds_7_6  + \
                     sanrentan_odds_7_8   + sanrentan_odds_7_9   + sanrentan_odds_7_10 + \
                     sanrentan_odds_7_11  + sanrentan_odds_7_12  + sanrentan_odds_7_13 + \
                     sanrentan_odds_7_14  + sanrentan_odds_7_15  + sanrentan_odds_7_16 + \
                     sanrentan_odds_7_17  + sanrentan_odds_7_18  + \
                     sanrentan_odds_8_1   + sanrentan_odds_8_2   + sanrentan_odds_8_3  + \
                     sanrentan_odds_8_4   + sanrentan_odds_8_5   + sanrentan_odds_8_6  + \
                     sanrentan_odds_8_7   + sanrentan_odds_8_9   + sanrentan_odds_8_10 + \
                     sanrentan_odds_8_11  + sanrentan_odds_8_12  + sanrentan_odds_8_13 + \
                     sanrentan_odds_8_14  + sanrentan_odds_8_15  + sanrentan_odds_8_16 + \
                     sanrentan_odds_8_17  + sanrentan_odds_8_18  + \
                     sanrentan_odds_9_1   + sanrentan_odds_9_2   + sanrentan_odds_9_3  + \
                     sanrentan_odds_9_4   + sanrentan_odds_9_5   + sanrentan_odds_9_6  + \
                     sanrentan_odds_9_7   + sanrentan_odds_9_8   + sanrentan_odds_9_10 + \
                     sanrentan_odds_9_11  + sanrentan_odds_9_12  + sanrentan_odds_9_13 + \
                     sanrentan_odds_9_14  + sanrentan_odds_9_15  + sanrentan_odds_9_16 + \
                     sanrentan_odds_9_17  + sanrentan_odds_9_18  + \
                     sanrentan_odds_10_1  + sanrentan_odds_10_2  + sanrentan_odds_10_3  + \
                     sanrentan_odds_10_4  + sanrentan_odds_10_5  + sanrentan_odds_10_6  + \
                     sanrentan_odds_10_7  + sanrentan_odds_10_8  + sanrentan_odds_10_9  + \
                     sanrentan_odds_10_11 + sanrentan_odds_10_12 + sanrentan_odds_10_13 + \
                     sanrentan_odds_10_14 + sanrentan_odds_10_15 + sanrentan_odds_10_16 + \
                     sanrentan_odds_10_17 + sanrentan_odds_10_18 + \
                     sanrentan_odds_11_1  + sanrentan_odds_11_2  + sanrentan_odds_11_3  + \
                     sanrentan_odds_11_4  + sanrentan_odds_11_5  + sanrentan_odds_11_6  + \
                     sanrentan_odds_11_7  + sanrentan_odds_11_8  + sanrentan_odds_11_9  + \
                     sanrentan_odds_11_10 + sanrentan_odds_11_12 + sanrentan_odds_11_13 + \
                     sanrentan_odds_11_14 + sanrentan_odds_11_15 + sanrentan_odds_11_16 + \
                     sanrentan_odds_11_17 + sanrentan_odds_11_18 + \
                     sanrentan_odds_12_1  + sanrentan_odds_12_2  + sanrentan_odds_12_3  + \
                     sanrentan_odds_12_4  + sanrentan_odds_12_5  + sanrentan_odds_12_6  + \
                     sanrentan_odds_12_7  + sanrentan_odds_12_8  + sanrentan_odds_12_9  + \
                     sanrentan_odds_12_10 + sanrentan_odds_12_11 + sanrentan_odds_12_13 + \
                     sanrentan_odds_12_14 + sanrentan_odds_12_15 + sanrentan_odds_12_16 + \
                     sanrentan_odds_12_17 + sanrentan_odds_12_18 + \
                     sanrentan_odds_13_1  + sanrentan_odds_13_2  + sanrentan_odds_13_3  + \
                     sanrentan_odds_13_4  + sanrentan_odds_13_5  + sanrentan_odds_13_6  + \
                     sanrentan_odds_13_7  + sanrentan_odds_13_8  + sanrentan_odds_13_9  + \
                     sanrentan_odds_13_10 + sanrentan_odds_13_11 + sanrentan_odds_13_12 + \
                     sanrentan_odds_13_14 + sanrentan_odds_13_15 + sanrentan_odds_13_16 + \
                     sanrentan_odds_13_17 + sanrentan_odds_13_18 + \
                     sanrentan_odds_14_1  + sanrentan_odds_14_2  + sanrentan_odds_14_3  + \
                     sanrentan_odds_14_4  + sanrentan_odds_14_5  + sanrentan_odds_14_6  + \
                     sanrentan_odds_14_7  + sanrentan_odds_14_8  + sanrentan_odds_14_9  + \
                     sanrentan_odds_14_10 + sanrentan_odds_14_11 + sanrentan_odds_14_12 + \
                     sanrentan_odds_14_13 + sanrentan_odds_14_15 + sanrentan_odds_14_16 + \
                     sanrentan_odds_14_17 + sanrentan_odds_14_18 + \
                     sanrentan_odds_15_1  + sanrentan_odds_15_2  + sanrentan_odds_15_3  + \
                     sanrentan_odds_15_4  + sanrentan_odds_15_5  + sanrentan_odds_15_6  + \
                     sanrentan_odds_15_7  + sanrentan_odds_15_8  + sanrentan_odds_15_9  + \
                     sanrentan_odds_15_10 + sanrentan_odds_15_11 + sanrentan_odds_15_12 + \
                     sanrentan_odds_15_13 + sanrentan_odds_15_14 + sanrentan_odds_15_16 + \
                     sanrentan_odds_15_17 + sanrentan_odds_15_18 + \
                     sanrentan_odds_16_1  + sanrentan_odds_16_2  + sanrentan_odds_16_3  + \
                     sanrentan_odds_16_4  + sanrentan_odds_16_5  + sanrentan_odds_16_6  + \
                     sanrentan_odds_16_7  + sanrentan_odds_16_8  + sanrentan_odds_16_9  + \
                     sanrentan_odds_16_10 + sanrentan_odds_16_11 + sanrentan_odds_16_12 + \
                     sanrentan_odds_16_13 + sanrentan_odds_16_14 + sanrentan_odds_16_15 + \
                     sanrentan_odds_16_17 + sanrentan_odds_16_18 + \
                     sanrentan_odds_17_1  + sanrentan_odds_17_2  + sanrentan_odds_17_3  + \
                     sanrentan_odds_17_4  + sanrentan_odds_17_5  + sanrentan_odds_17_6  + \
                     sanrentan_odds_17_7  + sanrentan_odds_17_8  + sanrentan_odds_17_9  + \
                     sanrentan_odds_17_10 + sanrentan_odds_17_11 + sanrentan_odds_17_12 + \
                     sanrentan_odds_17_13 + sanrentan_odds_17_14 + sanrentan_odds_17_15 + \
                     sanrentan_odds_17_16 + sanrentan_odds_17_18 + \
                     sanrentan_odds_18_1  + sanrentan_odds_18_2  + sanrentan_odds_18_3  + \
                     sanrentan_odds_18_4  + sanrentan_odds_18_5  + sanrentan_odds_18_6  + \
                     sanrentan_odds_18_7  + sanrentan_odds_18_8  + sanrentan_odds_18_9  + \
                     sanrentan_odds_18_10 + sanrentan_odds_18_11 + sanrentan_odds_18_12 + \
                     sanrentan_odds_18_13 + sanrentan_odds_18_14 + sanrentan_odds_18_15 + \
                     sanrentan_odds_18_16 + sanrentan_odds_18_17
    sanrentan_odds = [1000000.0 if i == '票数なし' else None if i == '' else i for i in sanrentan_odds]

    # データフレーム化
    sanrentan_odds_df = pd.DataFrame({
        'umaban_1': umaban_1,
        'umaban_2': umaban_2,
        'umaban_3': umaban_3,
        'sanrentan_odds': sanrentan_odds
    })

    sanrentan_odds_df = sanrentan_odds_df.query('umaban_1 != umaban_2 & \
        umaban_1 != umaban_3 & umaban_2 != umaban_3').dropna()
    sanrentan_odds_df.reset_index(drop=True, inplace=True)

    return(sanrentan_odds_df)
