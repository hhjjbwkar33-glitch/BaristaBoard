import pandas as pd
from datetime import datetime, timedelta

df = pd.read_excel('baristaboard_input_20260515.xlsx')
print(df)

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

def generate_break(start, end):
    work_time = datetime.strptime(end, '%H:%M:%S') - datetime.strptime(start, '%H:%M:%S')
    if work_time < timedelta(hours=5):
        break_time = 0
        return None
    elif work_time < timedelta(hours=6):
        break_time = 15
        
    elif work_time < timedelta(hours=7):
        break_time =  30
        
    else:
        break_time = 60
        
    start_break = datetime.strptime(start, '%H:%M:%S') + timedelta(hours=3)
    end_break = start_break + timedelta(minutes=break_time)

    return (start_break, end_break)

schedule = {}
break_schedule = {}
for index, row in df.iterrows():
    name = row['名前']
    start = row['開始時間']
    end = row['終了時間']

    slots = split_shift(start, end)
    for slot in slots:
        if slot in schedule:
            schedule[slot].append(name)
        else:
            schedule[slot] = [name]

    result = generate_break(start, end)
    if result is not None:
        if name in break_schedule:
            break_schedule[name].append(result)
        else:
            break_schedule[name] = [result]


positions = ['バリスタ', 'キャッシャー', 'フロア']
role_limit = 4
last_position = {}
streak_count = {}
new_schedule = {}
worked_minutes = {} #各従業員の累積勤務時間
break_remaining = {} #休憩の残り時間(分単位)

position_limit = {
    '通常': {'バリスタ': 1, 'キャッシャー': 1, 'フロア': 1},
    'ピーク': {'バリスタ': 2, 'キャッシャー': 1, 'フロア': 1}
}

def is_peak(slot):
    peak_slots = {'12:00-12:30', '12:30-13:00', '13:00-13:30', '13:30-14:00'}
    return slot in peak_slots

def required_positions(num_people):
    if num_people == 2:
        return ['バリスタ', 'キャッシャー']
    else:
        return ['バリスタ', 'キャッシャー', 'フロア']

def can_assign(name, position, current_limit, used_count):
    if used_count.get(position, 0) >= current_limit[position]:
        return False
    if last_position.get(name) == position:
        if streak_count.get(name, 0) >= role_limit:
            return False
    return True

def check_shortage(slot, assigned_list, num_people, peak):
    # ピーク・通常・人数に応じて required を決定
    if peak:
        required = {'バリスタ': 2, 'キャッシャー': 1, 'フロア': 1}
    elif num_people == 2:
        required = {'バリスタ': 1, 'キャッシャー': 1}
    else:
        required = {'バリスタ': 1, 'キャッシャー': 1, }

    actual = {}
    for name, position in assigned_list:
        actual[position] = actual.get(position, 0) + 1

    shortages = []
    for position, required_count in required.items():
        actual_count = actual.get(position, 0)
        if actual_count < required_count:
            shortages.append({
                'slot': slot,
                'position': position,
                'required': required_count,
                'actual': actual_count,
                'shortage': required_count - actual_count
            })
            print(
                f"⚠ 欠員発生: {slot}"
                f" | ポジション: {position}"
                f" | 不足: {required_count - actual_count}"
            )
    return shortages
# =====================

for slot, names in schedule.items():
    assigned = []
    assigned_people = set()
    num_people = len(names)
    required = required_positions(num_people)
    current_limit = position_limit['ピーク'] if is_peak(slot) else position_limit['通常']
    used_count = {p: 0 for p in current_limit}

    for position in required:
        for name in names:
            if name in assigned_people:
                continue
            if name in break_remaining:
                continue
            if name not in last_position:
                last_position[name] = position
                streak_count[name] = 0
            if can_assign(name, position, current_limit, used_count):
                if last_position[name] == position:
                    streak_count[name] += 1
                else:
                    streak_count[name] = 1
                last_position[name] = position
                used_count[position] += 1
                assigned.append((name, position))
                assigned_people.add(name)
                break

    new_schedule[slot] = assigned

#休憩ルール追加
    for name in names:  
        if name in break_remaining:
            print(f"  → 休憩中処理: {name} {break_remaining[name]} - 30")
            break_remaining[name] = break_remaining.get(name,0) - 30
            if break_remaining[name] <= 0:
                del break_remaining[name]
    

        else:
            worked_minutes[name] = worked_minutes.get(name, 0) + 30
            if worked_minutes[name] >= 360:
                break_remaining[name] = 45
                worked_minutes[name] = 0
    


records = []
shortage_log = []  # 欠員ログをまとめて保持

for slot, assigned in new_schedule.items():
    for name, position in assigned:
        records.append({'時間': slot, '名前': name, 'ポジション': position})

    
    num_people = len(schedule.get(slot, []))
    peak = is_peak(slot)
    shortages = check_shortage(slot, assigned, num_people, peak)
    shortage_log.extend(shortages)
    

result_df = pd.DataFrame(records)
print(result_df)

# 欠員サマリー表示
if shortage_log:
    print("\n===== 欠員サマリー =====")
    shortage_df = pd.DataFrame(shortage_log)
    print(shortage_df.to_string(index=False))
else:
    print("\n 欠員なし")

# Excel出力（シフト + 欠員ログを別シートに）
with pd.ExcelWriter('split_30minute_result.xlsx') as writer:
    result_df.to_excel(writer, sheet_name='シフト', index=False)
    if shortage_log:
        pd.DataFrame(shortage_log).to_excel(writer, sheet_name='欠員ログ', index=False)