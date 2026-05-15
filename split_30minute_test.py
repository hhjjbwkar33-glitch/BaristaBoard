import pandas as pd
#Excelファイルの読み込み
df = pd.read_excel('shift_test2.xlsx')
print(df)

from datetime import datetime, timedelta

def split_shift(start, end):
    
    start = str(start)
    end = str(end)
    
    split_time = []
    
    start_time = datetime.strptime(start, '%H:%M:%S')
    end_time = datetime.strptime(end, '%H:%M:%S')

    current_time = start_time
    
    while current_time < end_time:
        next_time = current_time + timedelta(minutes=30)
   
        time_interval = current_time.strftime('%H:%M') + '-' + next_time.strftime('%H:%M')
        split_time.append(time_interval)

        current_time = next_time
    return split_time


#時間→人への変換
schedule = {}

for index, row in df.iterrows():
    name = row['名前']
    start = row['開始時間']
    end = row['終了時間']
    
    #30分単位に分解
    slots = split_shift(start, end)
    
    #scheduleに格納
    for slot in slots:
        if slot in schedule: #時間帯と人を結びつける辞書にとって初めて見るかもしれないから条件分岐する
            schedule[slot].append(name)
        else:
            schedule[slot] = [name]


#ポジションの割り振り
positions = ['バリスタ', 'キャッシャー', 'フロア']

role_limit = 4 #同じポジションにつけのは最大2時間 = 30分x4

last_position = {} #この人は今どのポジションをやっているのか

streak_count = {}  #そのポジションを何スロット連続でやっているか

new_schedule = {}

position_limit = {
    '通常':{
        'バリスタ':1,
        'キャッシャー':1,
        'フロア':1

    },
    'ピーク':{
        'バリスタ':2,
        'キャッシャー':1,
        'フロア':1
    }
}
def is_peak(slot):
    peak_slots = {
        '12:00-12:30',
        '12:30-13:00',
        '13:00-13:30',
        '13:30-14:00',
    }
    return slot in peak_slots

def required_positions(num_people):
    if num_people == 2:
        return ['バリスタ', 'キャッシャー']
    else:
        return['バリスタ', 'キャッシャー', 'フロア']

#ポジションに入れていいかを判定
def can_assign(name, position, current_limit, used_count):

    #そのポジションが満員ならFalseを返す
    if used_count.get(position, 0) >= current_limit[position]:
        return False
    
    #前回と同じポジションかを確認
    if last_position.get(name) == position:
        #連続可能制限を超えるならFalseを返す
        if streak_count.get(name, 0) >= role_limit:
            return False
    
    return True

for slot, names in schedule.items():
    assigned = []

    assigned_people = set()

    num_people = len(names)
    required = required_positions(num_people)
    #三項演算子 A if 条件 else B → 「条件がTrueならA, FalseならB」
    current_limit = position_limit['ピーク'] if is_peak(slot) else position_limit['通常']
    #どのポジションが何人使われたかを初期化
    #辞書内包表記 {キー:値 for 要素 in リスト}
    used_count = {p:0 for p in current_limit}

    for position in required:

        for name in names:

        

            #すでにポジション配置済みならスキップ
            if name in assigned_people:
                continue

            #初めて出てきた人を初期化する
            if name not in last_position:
                last_position[name] = position
                streak_count[name] = 0
        
            #配置可能かをチェック
            if can_assign(name, position, current_limit, used_count):

                #前回と同じポジションなら連続回数（streak_count）追加
                if last_position[name] == position:
                    streak_count[name] += 1
                else:
                    streak_count[name] = 1
            
            #配置ポジション更新
            last_position[name] = position
            
            #使用人数（used_count）更新
            used_count[position] += 1

            #配置記録
            assigned.append((name, position))

            #この人は配置済み
            assigned_people.add(name)

            break
    new_schedule[slot] = assigned

#DataFrame化
records = []

for slot, assigned in new_schedule.items(): #new_scheduleの中身を1個ずつ取り出す

    for name, position in assigned: #assignedに登録されている一人一人の名前とポジションを取り出す

        records.append({
            '時間':slot,
            '名前':name,
            'ポジション':position
        })

result_df = pd.DataFrame(records)
print(result_df)

#Excel出力
result_df.to_excel(
    'split_30minute_result.xlsx',
    index=False

)