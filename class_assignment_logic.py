# -*- coding: utf-8 -*-
"""
학급 개편 추진 계획에 따른 반 배정 로직

원칙:
  가. 학년 학생 수를 4학급으로 나누고 반별 학생 수 결정
  나. 학년의 남녀 학생 수를 구분하여 균등하게 배정
"""


def calculate_class_distribution(total_students, male_count, female_count):
    """
    총 학생 수와 남녀 학생 수를 바탕으로 A, B, C, D반의 목표 인원을 계산합니다.
    
    원칙:
    - 남학생 수의 차이는 1명 이내
    - 여학생 수의 차이는 1명 이내
    - 총 학생 수의 차이는 1명 이내
    
    Args:
        total_students: 총 학생 수
        male_count: 남학생 수
        female_count: 여학생 수
    
    Returns:
        dict: 각 반의 목표 인원 정보
        {
            'total': {'A': int, 'B': int, 'C': int, 'D': int},
            'male': {'A': int, 'B': int, 'C': int, 'D': int},
            'female': {'A': int, 'B': int, 'C': int, 'D': int}
        }
    """
    # 검증
    if total_students != male_count + female_count:
        raise ValueError(f"총 학생 수({total_students})와 남녀 합계({male_count + female_count})가 일치하지 않습니다.")
    
    class_names = ['A', 'B', 'C', 'D']
    
    # 1. 남학생을 균등하게 배분 (차이 1명 이내)
    male_distribution = distribute_equally(male_count, class_names)
    
    # 2. 여학생을 균등하게 배분 (차이 1명 이내)
    female_distribution = distribute_equally(female_count, class_names)
    
    # 3. 각 반의 총원 = 남학생 + 여학생
    total_distribution = {
        name: male_distribution[name] + female_distribution[name]
        for name in class_names
    }
    
    # 검증: 각 반의 차이가 1명 이내인지 확인
    male_values = list(male_distribution.values())
    female_values = list(female_distribution.values())
    total_values = list(total_distribution.values())
    
    male_diff = max(male_values) - min(male_values)
    female_diff = max(female_values) - min(female_values)
    total_diff = max(total_values) - min(total_values)
    
    if male_diff > 1:
        raise ValueError(f"남학생 배분 차이가 1명을 초과합니다: {male_diff}명")
    if female_diff > 1:
        raise ValueError(f"여학생 배분 차이가 1명을 초과합니다: {female_diff}명")
    if total_diff > 1:
        raise ValueError(f"총 학생 수 배분 차이가 1명을 초과합니다: {total_diff}명")
    
    return {
        'total': total_distribution,
        'male': male_distribution,
        'female': female_distribution
    }


def distribute_equally(total_count, class_names):
    """
    총 인원을 여러 반에 균등 분배합니다.
    - 기본 몫: total_count // n
    - 나머지: A→B→C→D 순으로 1명씩 추가
    """
    num_classes = len(class_names)
    base = total_count // num_classes
    remainder = total_count % num_classes
    
    distribution = {name: base for name in class_names}
    for i in range(remainder):
        distribution[class_names[i]] += 1
    return distribution


def calculate_transfer_plan(previous_class_counts, target_distribution):
    """
    각 이전학반에서 목표 반으로 보낼 학생 수를 계산합니다.
    각 이전학반에서 보내는 총 인원(남+여)의 차이가 1명 이내가 되도록 하면서,
    전체적으로 목표 반별 필요 인원을 맞춥니다.
    
    전략:
    1. 각 이전학반의 남학생을 균등하게 분배 (4개 반에 균등)
    2. 각 이전학반의 여학생을 분배하되, 각 반의 총 인원(남+여)이 균등하도록
    3. 전체적으로 목표 반별 필요 인원을 맞추기 위해 반복 조정
    
    Args:
        previous_class_counts: 각 이전학반의 남녀 학생 수
            {
                1: {'male': int, 'female': int},
                2: {'male': int, 'female': int},
                3: {'male': int, 'female': int},
                4: {'male': int, 'female': int}
            }
        target_distribution: 목표 반별 남녀 인원 (calculate_class_distribution의 결과)
            {
                'male': {'A': int, 'B': int, 'C': int, 'D': int},
                'female': {'A': int, 'B': int, 'C': int, 'D': int}
            }
    
    Returns:
        dict: 각 이전학반에서 각 목표 반으로 보낼 학생 수
            {
                1: {'male': {'A': int, 'B': int, 'C': int, 'D': int},
                    'female': {'A': int, 'B': int, 'C': int, 'D': int}},
                2: {...},
                3: {...},
                4: {...}
            }
    """
    target_classes = ['A', 'B', 'C', 'D']
    previous_classes = [1, 2, 3, 4]
    
    # 목표 반별 필요 인원
    target_male_needs = {target: target_distribution['male'][target] for target in target_classes}
    target_female_needs = {target: target_distribution['female'][target] for target in target_classes}
    
    # 결과 초기화
    transfer_plan = {}
    for prev_class in previous_classes:
        transfer_plan[prev_class] = {
            'male': {target: 0 for target in target_classes},
            'female': {target: 0 for target in target_classes}
        }
    
    # 전체적으로 목표를 맞추는 방식으로 분배
    # 각 목표 반별로 필요한 남학생 수와 여학생 수를 각 이전학반에서 비례 분배
    
    # 각 목표 반별로 필요한 총 남학생 수와 여학생 수
    total_male_needed = sum(target_male_needs.values())
    total_female_needed = sum(target_female_needs.values())
    
    # 각 이전학반별로 처리
    # 먼저 각 이전학반에서 균등하게 분배 (숫자는 고정, 순서만 조정 가능)
    for prev_class in previous_classes:
        male_count = previous_class_counts[prev_class]['male']
        female_count = previous_class_counts[prev_class]['female']
        total_count = male_count + female_count
        
        if total_count == 0:
            continue
        
        # 남학생을 균등하게 분배 (4개 반에 균등)
        male_distribution = distribute_equally(male_count, target_classes)
        
        # 여학생을 균등하게 분배 (4개 반에 균등)
        female_distribution = distribute_equally(female_count, target_classes)
        
        transfer_plan[prev_class]['male'] = male_distribution
        transfer_plan[prev_class]['female'] = female_distribution
    
    # 전체 목표 분배를 맞추기 위해 조정
    # 각 이전학반에서 보내는 학생 수의 값은 고정하고, 순서만 바꿔서 전체 목표를 맞춤
    adjust_to_match_targets_preserving_distribution(transfer_plan, previous_classes, target_classes,
                          target_male_needs, target_female_needs)
    
    return transfer_plan


def adjust_to_match_targets_preserving_distribution(transfer_plan, previous_classes, target_classes,
                            target_male_needs, target_female_needs):
    """
    전체 목표를 맞추기 위해 조정하되, 각 이전학반에서 보내는 총 인원의 차이는 1명 이내로 유지.
    각 이전학반에서 보내는 학생 수의 패턴을 조정해서 전체 목표를 맞춤.
    """
    max_iterations = 200
    for iteration in range(max_iterations):
        # 각 목표 반으로 모이는 학생 수 재계산
        actual_male = {target: sum(transfer_plan[prev]['male'][target] for prev in previous_classes) 
                       for target in target_classes}
        actual_female = {target: sum(transfer_plan[prev]['female'][target] for prev in previous_classes) 
                         for target in target_classes}
        
        # 차이 계산
        male_diff = {target: target_male_needs[target] - actual_male[target] 
                     for target in target_classes}
        female_diff = {target: target_female_needs[target] - actual_female[target] 
                       for target in target_classes}
        
        # 모든 차이가 0이면 완료
        if all(d == 0 for d in male_diff.values()) and all(d == 0 for d in female_diff.values()):
            break
        
        improved = False
        
        # 각 이전학반에서 보내는 학생 수를 조정해서 전체 목표를 맞춤
        # 각 이전학반에서 보내는 총 인원의 차이는 1명 이내로 유지
        for prev_class in previous_classes:
            prev_male_total = sum(transfer_plan[prev_class]['male'].values())
            prev_female_total = sum(transfer_plan[prev_class]['female'].values())
            
            # 남학생 조정: 패턴을 조정해서 전체 목표에 맞춤
            # 단, 각 이전학반에서 보내는 학생 수의 차이는 1명 이내로 유지
            for target1 in target_classes:
                for target2 in target_classes:
                    if target1 == target2:
                        continue
                    
                    # target1에서 target2로 1명 이동
                    if transfer_plan[prev_class]['male'][target1] > 0:
                        # 이동 전 각 이전학반에서 보내는 학생 수의 차이 확인
                        before_male_vals = list(transfer_plan[prev_class]['male'].values())
                        before_male_diff = max(before_male_vals) - min(before_male_vals)
                        
                        # 이동
                        transfer_plan[prev_class]['male'][target1] -= 1
                        transfer_plan[prev_class]['male'][target2] += 1
                        
                        # 이동 후 각 이전학반에서 보내는 학생 수의 차이 확인
                        after_male_vals = list(transfer_plan[prev_class]['male'].values())
                        after_male_diff = max(after_male_vals) - min(after_male_vals)
                        
                        # 차이가 1명 이내인지 확인
                        if after_male_diff > 1:
                            # 되돌리기
                            transfer_plan[prev_class]['male'][target1] += 1
                            transfer_plan[prev_class]['male'][target2] -= 1
                            continue
                        
                        # 이동 전 차이
                        before_score = sum(abs(d) for d in male_diff.values())
                        
                        # 이동 후 차이 재계산
                        test_actual_male = {target: sum(transfer_plan[prev]['male'][target] for prev in previous_classes) 
                                           for target in target_classes}
                        test_male_diff = {target: target_male_needs[target] - test_actual_male[target] 
                                         for target in target_classes}
                        after_score = sum(abs(d) for d in test_male_diff.values())
                        
                        # 개선되었으면 유지, 아니면 되돌림
                        if after_score < before_score:
                            male_diff = test_male_diff
                            actual_male = test_actual_male
                            improved = True
                            break
                        else:
                            # 되돌리기
                            transfer_plan[prev_class]['male'][target1] += 1
                            transfer_plan[prev_class]['male'][target2] -= 1
                
                if improved:
                    break
            
            if improved:
                continue
            
            # 여학생 조정: 패턴을 조정해서 전체 목표에 맞춤
            # 단, 각 이전학반에서 보내는 학생 수의 차이는 1명 이내로 유지
            for target1 in target_classes:
                for target2 in target_classes:
                    if target1 == target2:
                        continue
                    
                    # target1에서 target2로 1명 이동
                    if transfer_plan[prev_class]['female'][target1] > 0:
                        # 이동 전 각 이전학반에서 보내는 학생 수의 차이 확인
                        before_female_vals = list(transfer_plan[prev_class]['female'].values())
                        before_female_diff = max(before_female_vals) - min(before_female_vals)
                        
                        # 이동
                        transfer_plan[prev_class]['female'][target1] -= 1
                        transfer_plan[prev_class]['female'][target2] += 1
                        
                        # 이동 후 각 이전학반에서 보내는 학생 수의 차이 확인
                        after_female_vals = list(transfer_plan[prev_class]['female'].values())
                        after_female_diff = max(after_female_vals) - min(after_female_vals)
                        
                        # 차이가 1명 이내인지 확인
                        if after_female_diff > 1:
                            # 되돌리기
                            transfer_plan[prev_class]['female'][target1] += 1
                            transfer_plan[prev_class]['female'][target2] -= 1
                            continue
                        
                        # 이동 전 차이
                        before_score = sum(abs(d) for d in female_diff.values())
                        
                        # 이동 후 차이 재계산
                        test_actual_female = {target: sum(transfer_plan[prev]['female'][target] for prev in previous_classes) 
                                            for target in target_classes}
                        test_female_diff = {target: target_female_needs[target] - test_actual_female[target] 
                                           for target in target_classes}
                        after_score = sum(abs(d) for d in test_female_diff.values())
                        
                        # 개선되었으면 유지, 아니면 되돌림
                        if after_score < before_score:
                            female_diff = test_female_diff
                            actual_female = test_actual_female
                            improved = True
                            break
                        else:
                            # 되돌리기
                            transfer_plan[prev_class]['female'][target1] += 1
                            transfer_plan[prev_class]['female'][target2] -= 1
                
                if improved:
                    break
        
        if not improved:
            break


def adjust_to_match_targets(transfer_plan, previous_classes, target_classes,
                            target_male_needs, target_female_needs,
                            male_diff, female_diff):
    """
    전체적으로 목표를 맞추기 위해 이동 계획을 조정합니다.
    각 이전학반에서 보내는 총 인원의 차이는 1명 이내로 유지합니다.
    """
    max_iterations = 100
    for iteration in range(max_iterations):
        # 각 목표 반으로 모이는 학생 수 재계산
        actual_male = {target: sum(transfer_plan[prev]['male'][target] for prev in previous_classes) 
                       for target in target_classes}
        actual_female = {target: sum(transfer_plan[prev]['female'][target] for prev in previous_classes) 
                         for target in target_classes}
        
        # 차이 계산
        male_diff = {target: target_male_needs[target] - actual_male[target] 
                     for target in target_classes}
        female_diff = {target: target_female_needs[target] - actual_female[target] 
                       for target in target_classes}
        
        # 모든 차이가 0이면 완료
        if all(d == 0 for d in male_diff.values()) and all(d == 0 for d in female_diff.values()):
            break
        
        improved = False
        
        # 각 이전학반에서 보내는 총 인원 계산 (균등 분배 기준)
        prev_totals = {}
        for prev_class in previous_classes:
            prev_totals[prev_class] = sum(transfer_plan[prev_class]['male'].values()) + sum(transfer_plan[prev_class]['female'].values())
        
        # 조정: 부족한 반에 더 보내고, 초과한 반에서 덜 보내기
        # 각 이전학반에서 보내는 총 인원의 차이는 1명 이내로 유지하면서
        for prev_class in previous_classes:
            # 남학생 조정
            for target in target_classes:
                if male_diff[target] > 0:
                    # 부족한 반에 더 보내기
                    # 다른 반에서 가져올 수 있는지 확인
                    for other_target in target_classes:
                        if other_target != target and male_diff[other_target] < 0:
                            # other_target에서 target으로 이동
                            if transfer_plan[prev_class]['male'][other_target] > 0:
                                # 이동 후 각 이전학반에서 보내는 총 인원의 차이 확인
                                # 이동 전후 총 인원은 변하지 않으므로 차이는 유지됨
                                # 하지만 각 반으로 보내는 인원의 균등성을 확인
                                before_male = list(transfer_plan[prev_class]['male'].values())
                                before_diff = max(before_male) - min(before_male)
                                
                                transfer_plan[prev_class]['male'][other_target] -= 1
                                transfer_plan[prev_class]['male'][target] += 1
                                
                                after_male = list(transfer_plan[prev_class]['male'].values())
                                after_diff = max(after_male) - min(after_male)
                                
                                # 균등성이 개선되거나 유지되는 경우만 허용
                                if after_diff <= before_diff + 1:  # 최대 1명 차이 증가 허용
                                    male_diff[target] -= 1
                                    male_diff[other_target] += 1
                                    improved = True
                                    break
                                else:
                                    # 되돌리기
                                    transfer_plan[prev_class]['male'][other_target] += 1
                                    transfer_plan[prev_class]['male'][target] -= 1
                    if improved:
                        break
            
            if improved:
                continue
            
            # 여학생 조정
            for target in target_classes:
                if female_diff[target] > 0:
                    # 부족한 반에 더 보내기
                    for other_target in target_classes:
                        if other_target != target and female_diff[other_target] < 0:
                            if transfer_plan[prev_class]['female'][other_target] > 0:
                                # 이동 전후 균등성 확인
                                before_female = list(transfer_plan[prev_class]['female'].values())
                                before_diff = max(before_female) - min(before_female)
                                
                                transfer_plan[prev_class]['female'][other_target] -= 1
                                transfer_plan[prev_class]['female'][target] += 1
                                
                                after_female = list(transfer_plan[prev_class]['female'].values())
                                after_diff = max(after_female) - min(after_female)
                                
                                # 균등성이 개선되거나 유지되는 경우만 허용
                                if after_diff <= before_diff + 1:  # 최대 1명 차이 증가 허용
                                    female_diff[target] -= 1
                                    female_diff[other_target] += 1
                                    improved = True
                                    break
                                else:
                                    # 되돌리기
                                    transfer_plan[prev_class]['female'][other_target] += 1
                                    transfer_plan[prev_class]['female'][target] -= 1
                    if improved:
                        break
        
        if not improved:
            break


def distribute_proportional_to_targets(student_count, target_needs, target_classes):
    """
    목표 반별 필요 인원에 비례하여 분배합니다.
    각 이전학반에서 보내는 총 인원의 차이는 1명 이내로 유지합니다.
    
    Args:
        student_count: 이전학반의 학생 수 (남 또는 여)
        target_needs: 각 목표 반의 필요 인원 {'A': int, 'B': int, ...}
        target_classes: 목표 반 리스트
    
    Returns:
        dict: 각 목표 반으로 보낼 학생 수
    """
    num_targets = len(target_classes)
    total_needed = sum(target_needs.values())
    
    if student_count == 0:
        return {target: 0 for target in target_classes}
    
    if total_needed == 0:
        # 목표가 없으면 균등 분배
        return distribute_from_previous_class(student_count, {}, target_classes)
    
    # 목표 반별 필요 비율 계산
    ratios = {}
    for target in target_classes:
        if total_needed > 0:
            ratios[target] = target_needs[target] / total_needed
        else:
            ratios[target] = 1.0 / num_targets
    
    # 비율에 따라 기본 배정
    distribution = {}
    allocated = 0
    
    for target in target_classes:
        base_count = int(student_count * ratios[target])
        distribution[target] = base_count
        allocated += base_count
    
    # 나머지 인원을 목표 필요 인원이 많은 순서로 배정
    remainder = student_count - allocated
    if remainder > 0:
        # 목표 필요 인원이 많은 순서로 정렬
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (target_needs[t], target_classes.index(t)),
            reverse=True
        )
        
        for target in sorted_targets:
            if remainder <= 0:
                break
            distribution[target] += 1
            remainder -= 1
    
    return distribution


def distribute_female_balanced(female_count, male_distribution, target_female_needs, target_classes):
    """
    여학생을 분배하되, 각 반의 총 인원(남+여)이 균등하도록 하면서,
    목표 반별 필요 인원도 고려합니다.
    """
    num_targets = len(target_classes)
    
    if female_count == 0:
        return {target: 0 for target in target_classes}
    
    # 각 반의 현재 총 인원(남학생만) 계산
    current_totals = {target: male_distribution[target] for target in target_classes}
    
    # 목표: 각 반의 총 인원(남+여)이 균등하도록
    total_students = sum(male_distribution.values()) + female_count
    target_avg = total_students / num_targets
    
    # 각 반에 필요한 여학생 수 계산 (목표 총 인원 - 현재 남학생 수)
    distribution = {}
    for target in target_classes:
        needed = max(0, int(target_avg) - current_totals[target])
        distribution[target] = needed
    
    allocated = sum(distribution.values())
    remainder = female_count - allocated
    
    # 나머지 처리: 총 인원이 적은 반부터 추가 배정
    if remainder > 0:
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (current_totals[t] + distribution[t], target_classes.index(t))
        )
        for target in sorted_targets:
            if remainder <= 0:
                break
            distribution[target] += 1
            remainder -= 1
    elif remainder < 0:
        # 초과한 경우 총 인원이 많은 반부터 차감
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (current_totals[t] + distribution[t], target_classes.index(t)),
            reverse=True
        )
        for target in sorted_targets:
            if remainder >= 0:
                break
            if distribution[target] > 0:
                distribution[target] -= 1
                remainder += 1
    
    return distribution




def distribute_with_targets_and_balance(student_count, target_needs, target_classes, prev_total):
    """
    목표 반별 필요 인원에 비례하여 분배하되,
    각 이전학반에서 보내는 총 인원의 차이가 1명 이내가 되도록 합니다.
    
    Args:
        student_count: 이전학반의 학생 수 (남 또는 여)
        target_needs: 각 목표 반의 필요 인원 {'A': int, 'B': int, ...}
        target_classes: 목표 반 리스트
        prev_total: 이전학반의 총 학생 수 (균등 분배를 위한 기준)
    
    Returns:
        dict: 각 목표 반으로 보낼 학생 수
    """
    num_targets = len(target_classes)
    total_needed = sum(target_needs.values())
    
    if student_count == 0:
        return {target: 0 for target in target_classes}
    
    if total_needed == 0:
        # 목표가 없으면 균등 분배
        return distribute_from_previous_class(student_count, {}, target_classes)
    
    # 목표 반별 필요 비율 계산
    ratios = {}
    for target in target_classes:
        if total_needed > 0:
            ratios[target] = target_needs[target] / total_needed
        else:
            ratios[target] = 1.0 / num_targets
    
    # 비율에 따라 기본 배정
    distribution = {}
    allocated = 0
    
    for target in target_classes:
        base_count = int(student_count * ratios[target])
        distribution[target] = base_count
        allocated += base_count
    
    # 나머지 인원을 목표 필요 인원이 많은 순서로 배정
    remainder = student_count - allocated
    if remainder > 0:
        # 목표 필요 인원이 많은 순서로 정렬
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (target_needs[t], target_classes.index(t)),
            reverse=True
        )
        
        for target in sorted_targets:
            if remainder <= 0:
                break
            distribution[target] += 1
            remainder -= 1
    
    return distribution


def distribute_female_to_balance_total(female_count, male_distribution, target_classes):
    """
    여학생을 분배하되, 각 반으로 보내는 총 인원(남+여)이 균등하도록 합니다.
    
    Args:
        female_count: 여학생 수
        male_distribution: 이미 분배된 남학생 수 {'A': int, 'B': int, ...}
        target_classes: 목표 반 리스트
    
    Returns:
        dict: 각 목표 반으로 보낼 여학생 수
    """
    num_targets = len(target_classes)
    
    if female_count == 0:
        return {target: 0 for target in target_classes}
    
    # 각 반의 현재 총 인원(남학생만) 계산
    current_totals = {target: male_distribution[target] for target in target_classes}
    
    # 목표: 각 반의 총 인원(남+여)이 균등하도록
    # 총 인원(남+여)의 평균 계산
    total_students = sum(male_distribution.values()) + female_count
    target_avg = total_students / num_targets
    
    # 각 반에 필요한 여학생 수 계산 (목표 총 인원 - 현재 남학생 수)
    needed_females = {}
    for target in target_classes:
        needed = max(0, int(target_avg) - current_totals[target])
        needed_females[target] = needed
    
    # 기본 배정
    distribution = needed_females.copy()
    allocated = sum(distribution.values())
    
    # 나머지 처리
    remainder = female_count - allocated
    if remainder > 0:
        # 총 인원이 적은 반부터 추가 배정
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (current_totals[t] + distribution[t], target_classes.index(t))
        )
        for target in sorted_targets:
            if remainder <= 0:
                break
            distribution[target] += 1
            remainder -= 1
    elif remainder < 0:
        # 초과한 경우 총 인원이 많은 반부터 차감
        sorted_targets = sorted(
            target_classes,
            key=lambda t: (current_totals[t] + distribution[t], target_classes.index(t)),
            reverse=True
        )
        for target in sorted_targets:
            if remainder >= 0:
                break
            if distribution[target] > 0:
                distribution[target] -= 1
                remainder += 1
    
    return distribution


def distribute_from_previous_class(student_count, target_needs, target_classes):
    """
    이전학반의 학생을 목표 반별 필요 인원에 맞춰 분배합니다.
    각 이전학반에서 보내는 총 인원의 차이가 1명 이내가 되도록 합니다.
    
    Args:
        student_count: 이전학반의 학생 수
        target_needs: 각 목표 반의 필요 인원 {'A': int, 'B': int, ...}
        target_classes: 목표 반 리스트
    
    Returns:
        dict: 각 목표 반으로 보낼 학생 수
    """
    num_targets = len(target_classes)
    
    if student_count == 0:
        return {target: 0 for target in target_classes}
    
    # 기본 몫 계산
    base = student_count // num_targets
    remainder = student_count % num_targets
    
    # 기본값으로 초기화
    distribution = {target: base for target in target_classes}
    
    # 나머지를 A→B→C→D 순서로 1명씩 추가
    for i in range(remainder):
        distribution[target_classes[i]] += 1
    
    return distribution


def print_transfer_plan(transfer_plan, previous_class_counts, target_distribution):
    """
    이동 계획을 보기 좋게 출력합니다.
    """
    print("=" * 70)
    print("각 이전학반에서 목표 반으로 보낼 학생 수")
    print("=" * 70)
    print()
    
    target_classes = ['A', 'B', 'C', 'D']
    previous_classes = [1, 2, 3, 4]
    
    # 각 이전학반별 출력
    for prev_class in previous_classes:
        male_count = previous_class_counts[prev_class]['male']
        female_count = previous_class_counts[prev_class]['female']
        total_count = male_count + female_count
        
        print(f"[{prev_class}반] (총 {total_count}명: 남 {male_count}명, 여 {female_count}명)")
        print("-" * 70)
        
        # 남학생 이동 계획
        male_transfers = transfer_plan[prev_class]['male']
        male_sum = sum(male_transfers.values())
        print(f"  남학생 이동: ", end="")
        for target in target_classes:
            print(f"{target}반 {male_transfers[target]}명", end="")
            if target != target_classes[-1]:
                print(", ", end="")
        print(f" (합계: {male_sum}명)")
        
        # 여학생 이동 계획
        female_transfers = transfer_plan[prev_class]['female']
        female_sum = sum(female_transfers.values())
        print(f"  여학생 이동: ", end="")
        for target in target_classes:
            print(f"{target}반 {female_transfers[target]}명", end="")
            if target != target_classes[-1]:
                print(", ", end="")
        print(f" (합계: {female_sum}명)")
        
        # 총 이동 인원
        total_transfer = male_sum + female_sum
        print(f"  총 이동 인원: {total_transfer}명")
        print()
    
    # 검증: 각 목표 반으로 모이는 학생 수
    print("=" * 70)
    print("검증: 각 목표 반으로 모이는 학생 수")
    print("=" * 70)
    print()
    
    print(f"{'목표반':<8} {'목표 남':<10} {'실제 남':<10} {'목표 여':<10} {'실제 여':<10} {'목표 총':<10} {'실제 총':<10}")
    print("-" * 70)
    
    for target in target_classes:
        target_male = target_distribution['male'][target]
        target_female = target_distribution['female'][target]
        target_total = target_distribution['total'][target]
        
        # 실제로 모이는 학생 수 계산
        actual_male = sum(transfer_plan[prev]['male'][target] for prev in previous_classes)
        actual_female = sum(transfer_plan[prev]['female'][target] for prev in previous_classes)
        actual_total = actual_male + actual_female
        
        male_match = "[OK]" if actual_male == target_male else "[ERROR]"
        female_match = "[OK]" if actual_female == target_female else "[ERROR]"
        total_match = "[OK]" if actual_total == target_total else "[ERROR]"
        
        print(f"{target:<8} {target_male:<10} {actual_male:<10}{male_match:<4} {target_female:<10} {actual_female:<10}{female_match:<4} {target_total:<10} {actual_total:<10}{total_match}")
    
    print()


def print_distribution_summary(distribution, total_students, male_count, female_count):
    """
    배정 결과를 보기 좋게 출력합니다.
    """
    print("=" * 70)
    print("반별 인원 배정 결과")
    print("=" * 70)
    print()
    
    print(f"전체 현황")
    print(f"  - 총 학생 수: {total_students}명")
    print(f"  - 남학생: {male_count}명")
    print(f"  - 여학생: {female_count}명")
    print()
    
    print("=" * 70)
    print("반별 목표 인원")
    print("=" * 70)
    print()
    
    total_dist = distribution['total']
    male_dist = distribution['male']
    female_dist = distribution['female']
    
    print(f"{'반':<6} {'총 인원':<10} {'남학생':<10} {'여학생':<10}")
    print("-" * 70)
    
    for class_name in ['A', 'B', 'C', 'D']:
        total = total_dist[class_name]
        male = male_dist[class_name]
        female = female_dist[class_name]
        print(f"{class_name:<6} {total:<10} {male:<10} {female:<10}")
    
    print("-" * 70)
    print(f"{'합계':<6} {sum(total_dist.values()):<10} {sum(male_dist.values()):<10} {sum(female_dist.values()):<10}")
    print()
    
    # 검증
    print("=" * 70)
    print("검증")
    print("=" * 70)
    total_sum = sum(total_dist.values())
    male_sum = sum(male_dist.values())
    female_sum = sum(female_dist.values())
    
    print(f"  - 총 인원 합계: {total_sum}명 (목표: {total_students}명) {'[OK]' if total_sum == total_students else '[ERROR]'}")
    print(f"  - 남학생 합계: {male_sum}명 (목표: {male_count}명) {'[OK]' if male_sum == male_count else '[ERROR]'}")
    print(f"  - 여학생 합계: {female_sum}명 (목표: {female_count}명) {'[OK]' if female_sum == female_count else '[ERROR]'}")
    print()


if __name__ == "__main__":
    # 현재 데이터로 테스트
    import pandas as pd
    
    # 학생자료 읽기
    df = pd.read_excel('학생자료.xlsx', sheet_name=0)
    
    # 남녀 학생 수 계산
    total_students = len(df)
    male_count = len(df[df['남녀'] == '남'])
    female_count = len(df[df['남녀'] == '여'])
    
    print("=" * 70)
    print("학급 개편 반 배정 로직 테스트")
    print("=" * 70)
    print()
    
    # 배정 계산
    distribution = calculate_class_distribution(total_students, male_count, female_count)
    
    # 결과 출력
    print_distribution_summary(distribution, total_students, male_count, female_count)
    
    # 각 이전학반의 남녀 학생 수 계산
    previous_class_counts = {}
    for prev_class in [1, 2, 3, 4]:
        prev_df = df[df['이전학반'] == prev_class]
        previous_class_counts[prev_class] = {
            'male': len(prev_df[prev_df['남녀'] == '남']),
            'female': len(prev_df[prev_df['남녀'] == '여'])
        }
    
    # 이동 계획 계산
    transfer_plan = calculate_transfer_plan(previous_class_counts, distribution)
    
    # 이동 계획 출력
    print_transfer_plan(transfer_plan, previous_class_counts, distribution)
    
    # 예시: 다른 학생 수로도 테스트
    print("\n" + "=" * 70)
    print("다른 학생 수 예시 테스트")
    print("=" * 70)
    print()
    
    # 예시 1: 128명, 남66명, 여62명
    print("예시 1: 총 128명 (남 66명, 여 62명)")
    print("-" * 70)
    dist1 = calculate_class_distribution(128, 66, 62)
    print_distribution_summary(dist1, 128, 66, 62)
    
    # 예시 이전학반 구성 (가정)
    example_previous = {
        1: {'male': 16, 'female': 14},
        2: {'male': 16, 'female': 14},
        3: {'male': 17, 'female': 15},
        4: {'male': 17, 'female': 15}
    }
    example_transfer = calculate_transfer_plan(example_previous, dist1)
    print_transfer_plan(example_transfer, example_previous, dist1)
    
    # 예시 2: 120명, 남60명, 여60명
    print("\n예시 2: 총 120명 (남 60명, 여 60명)")
    print("-" * 70)
    dist2 = calculate_class_distribution(120, 60, 60)
    print_distribution_summary(dist2, 120, 60, 60)
