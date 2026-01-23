# -*- coding: utf-8 -*-
"""
학급 개편 반 배정 로직 (v2) — 3반 편성 지원

기존 class_assignment_logic을 그대로 사용하고,
3반 편성 시 "선생님 연속 지도 배제" 규칙을 추가합니다.

- 4반 편성: 1,2,3,4반 → A,B,C,D 자유 배정 (기존과 동일)
- 3반 편성: 1→B,C,D / 2→A,C,D / 3→A,B,D / 4→A,B,C (해당 번호 반 제외)
"""

from class_assignment_logic import (
    distribute_equally,
    calculate_class_distribution,
    print_distribution_summary,
    print_transfer_plan,
)

# 3반 편성 시 이전학반별 배정 가능한 목표 반 (현재 반 담당 선생님이 다음 학년에 맡지 않도록)
# 1반→A 제외, 2반→B 제외, 3반→C 제외, 4반→A 제외
ALLOWED_TARGETS_3반 = {
    1: ['B', 'C', 'D'],
    2: ['A', 'C', 'D'],
    3: ['A', 'B', 'D'],
    4: ['B', 'C', 'D'],
}

TARGET_CLASSES = ['A', 'B', 'C', 'D']
PREVIOUS_CLASSES = [1, 2, 3, 4]


def calculate_transfer_plan_3반(previous_class_counts, target_distribution):
    """
    3반 편성용 이동 계획.
    각 이전학반은 해당 번호 반(A/B/C/D)으로 보낼 수 없고,
    전체 목표 인원·남녀 균형(차이 1명 이내)은 유지합니다.
    """
    target_male_needs = {t: target_distribution['male'][t] for t in TARGET_CLASSES}
    target_female_needs = {t: target_distribution['female'][t] for t in TARGET_CLASSES}

    transfer_plan = {}
    for prev in PREVIOUS_CLASSES:
        transfer_plan[prev] = {
            'male': {t: 0 for t in TARGET_CLASSES},
            'female': {t: 0 for t in TARGET_CLASSES},
        }

    for prev_class in PREVIOUS_CLASSES:
        male_count = previous_class_counts[prev_class]['male']
        female_count = previous_class_counts[prev_class]['female']
        allowed = ALLOWED_TARGETS_3반[prev_class]

        if male_count + female_count == 0:
            continue

        male_dist = distribute_equally(male_count, allowed)
        female_dist = distribute_equally(female_count, allowed)

        for t in TARGET_CLASSES:
            transfer_plan[prev_class]['male'][t] = male_dist.get(t, 0)
            transfer_plan[prev_class]['female'][t] = female_dist.get(t, 0)

    adjust_to_match_targets_3반(
        transfer_plan, PREVIOUS_CLASSES, TARGET_CLASSES,
        target_male_needs, target_female_needs,
    )
    return transfer_plan


def adjust_to_match_targets_3반(transfer_plan, previous_classes, target_classes,
                                target_male_needs, target_female_needs):
    """
    전체 목표를 맞추기 위해 조정하되,
    - 각 이전학반에서 보내는 총 인원 차이 1명 이내 유지
    - 3반 편성: 이동은 해당 이전학반의 allowed 대상 반끼리만 가능
    """
    max_iterations = 200
    for _ in range(max_iterations):
        actual_male = {t: sum(transfer_plan[p]['male'][t] for p in previous_classes) for t in target_classes}
        actual_female = {t: sum(transfer_plan[p]['female'][t] for p in previous_classes) for t in target_classes}

        male_diff = {t: target_male_needs[t] - actual_male[t] for t in target_classes}
        female_diff = {t: target_female_needs[t] - actual_female[t] for t in target_classes}

        if all(d == 0 for d in male_diff.values()) and all(d == 0 for d in female_diff.values()):
            break

        improved = False

        for prev_class in previous_classes:
            allowed = list(ALLOWED_TARGETS_3반[prev_class])

            # 남학생 조정 (allowed 내에서만 target1 <-> target2 이동)
            for t1 in allowed:
                others = [t for t in allowed if t != t1]
                for t2 in others:
                    if transfer_plan[prev_class]['male'][t1] <= 0:
                        continue

                    before_vals = [transfer_plan[prev_class]['male'][t] for t in allowed]
                    before_d = max(before_vals) - min(before_vals)

                    transfer_plan[prev_class]['male'][t1] -= 1
                    transfer_plan[prev_class]['male'][t2] += 1

                    after_vals = [transfer_plan[prev_class]['male'][t] for t in allowed]
                    after_d = max(after_vals) - min(after_vals)
                    if after_d > 1:
                        transfer_plan[prev_class]['male'][t1] += 1
                        transfer_plan[prev_class]['male'][t2] -= 1
                        continue

                    before_score = sum(abs(male_diff[t]) for t in target_classes)
                    test_male = {t: sum(transfer_plan[p]['male'][t] for p in previous_classes) for t in target_classes}
                    test_diff = {t: target_male_needs[t] - test_male[t] for t in target_classes}
                    after_score = sum(abs(test_diff[t]) for t in target_classes)

                    if after_score < before_score:
                        male_diff.update(test_diff)
                        improved = True
                        break
                    transfer_plan[prev_class]['male'][t1] += 1
                    transfer_plan[prev_class]['male'][t2] -= 1
                if improved:
                    break

            if improved:
                continue

            # 여학생 조정 (allowed 내에서만)
            for t1 in allowed:
                others = [t for t in allowed if t != t1]
                for t2 in others:
                    if transfer_plan[prev_class]['female'][t1] <= 0:
                        continue

                    before_vals = [transfer_plan[prev_class]['female'][t] for t in allowed]
                    before_d = max(before_vals) - min(before_vals)

                    transfer_plan[prev_class]['female'][t1] -= 1
                    transfer_plan[prev_class]['female'][t2] += 1

                    after_vals = [transfer_plan[prev_class]['female'][t] for t in allowed]
                    after_d = max(after_vals) - min(after_vals)
                    if after_d > 1:
                        transfer_plan[prev_class]['female'][t1] += 1
                        transfer_plan[prev_class]['female'][t2] -= 1
                        continue

                    before_score = sum(abs(female_diff[t]) for t in target_classes)
                    test_female = {t: sum(transfer_plan[p]['female'][t] for p in previous_classes) for t in target_classes}
                    test_diff = {t: target_female_needs[t] - test_female[t] for t in target_classes}
                    after_score = sum(abs(test_diff[t]) for t in target_classes)

                    if after_score < before_score:
                        female_diff.update(test_diff)
                        improved = True
                        break
                    transfer_plan[prev_class]['female'][t1] += 1
                    transfer_plan[prev_class]['female'][t2] -= 1
                if improved:
                    break

        if not improved:
            break
