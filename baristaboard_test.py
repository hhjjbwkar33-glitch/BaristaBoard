import pandas as pd
#Excelファイルの読み込み
df = pd.read_excel('shift_test.xlsx')
print(df)

#時間分解関数の作成
def split_shift(start, end):
    split_time = [] #分割した時間を保存するリスト
    start_time = start.hour #勤務開始時間を取り出す
    end_time = end.hour #勤務終了時間を取り出す
    
    for hour in range(start_time, end_time):
        split_time.append(f'{hour}-{hour+1}')

    return split_time

#勤務開始・終了時間の取り出し
for index, row in df.iterrows():

    slots = split_shift(
        row['開始時間'],
        row['終了時間']

    )

#時間→人への変換
schedule = {}
for index, row in df.iterrows():
    name = row['名前']
    slots = split_shift(
        row['開始時間'],
        row['終了時間']

    )
    for slot in slots:
        if slot in schedule: #時間帯と人を結びつける辞書にとって初めて見るかもしれないから条件分岐する
            schedule[slot].append(name)
        else:
            schedule[slot] = [name]
print(schedule)

#時間帯-人のデータを整形
for slot, names in schedule.items(): #schedule辞書内のキー（時間帯）と値（人名）をセットで取り出し、slotとnameにそれぞれ代入
    print(slot, ':', names)

#ポジションの割り振り
positions = ['バリスタ', 'キャッシャー', 'フロア']
new_schedule = {}

for slot, names in schedule.items(): 

    assigned = [] #時間帯とポジションの結果を入れる

    for i, name in enumerate(names): #名前とその名前が何番目かを同時に取得し変数に代入
        position = positions[i % len(positions)] #ポジションを順番にローテーション
        assigned.append((name, position))
    new_schedule[slot] = assigned
    
    print(slot)
    for name, position in assigned:
        print(name, position)