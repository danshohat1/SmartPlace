"""University-proposing Gale–Shapley with weights, ties, and placement explanations."""

from typing import Dict, List, Tuple
from SmartPlace.models import Student, University

def get_rank(student: Student, university_name: str) -> int:
    """Return the tier index of `university_name` in student.preferences."""
    for i, tier in enumerate(student.preferences):
        if isinstance(tier, str):
            if tier == university_name:
                return i
        else:  # list
            if university_name in tier:
                return i
    return len(student.preferences)


def weighted_gale_shapley(
    students: Dict[str, Student],
    universities: Dict[str, University],
    gamma: float = 1.0
) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Run weighted university-proposing GS; returns matches and plain-English reasons."""
    n = len(universities)
    matches: Dict[str, str] = {}
    reasons: Dict[str, str] = {s: "" for s in students}

    def ordinal(i: int) -> str:
        i1 = i + 1
        if 10 < i1 % 100 < 14:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(i1 % 10, "th")
        return f"{i1}{suffix}"

    def components(s: Student, u: University):
        r = get_rank(s, u.name)
        stu_comp = s.voice * (n - r)
        uni_comp = gamma * u.power
        return r, stu_comp, uni_comp, stu_comp + uni_comp

    def dominant_factor(delta_pref: float, delta_power: float) -> str:
        # Decide which qualitative factor mainly changed
        if abs(delta_pref) > abs(delta_power) * 1.2:
            return "preference"
        if abs(delta_power) > abs(delta_pref) * 1.2:
            return "power"
        return "mixed"

    free_unis: List[University] = [u for u in universities.values() if u.has_free_slot()]

    while free_unis:
        uni = free_unis.pop(0)
        cand = uni.next_candidate()
        if not cand or cand not in students:
            continue
        stu = students[cand]

        r_new, stu_comp_new, uni_comp_new, total_new = components(stu, uni)

        if stu.match is None:
            stu.match = uni.name
            uni.accepted.append(cand)
            matches[cand] = uni.name
            reasons[cand] = (
                f"Matched with {uni.name} because it was the first suitable offer received. "
                f"The student had no prior commitment."
            )
        else:
            current = universities[stu.match]
            r_old, stu_comp_old, uni_comp_old, total_old = components(stu, current)
            if total_new > total_old:
                # Determine qualitative reason
                pref_improved = r_new < r_old
                power_improved = uni_comp_new > uni_comp_old
                factor = dominant_factor(stu_comp_new - stu_comp_old, uni_comp_new - uni_comp_old)
                if pref_improved and not power_improved:
                    reason_core = (
                        f"Switched from {current.name} to {uni.name} because {uni.name} is higher on "
                        f"the student's preference list (from {ordinal(r_old)} to {ordinal(r_new)})."
                    )
                elif power_improved and not pref_improved:
                    reason_core = (
                        f"Switched from {current.name} to {uni.name} even though preference rank did not improve, "
                        f"because {uni.name} carries stronger institutional weight."
                    )
                elif pref_improved and power_improved:
                    if factor == "preference":
                        reason_core = (
                            f"Switched from {current.name} to {uni.name}; main reason: higher preference ("
                            f"{ordinal(r_old)} → {ordinal(r_new)}) plus stronger university weight."
                        )
                    elif factor == "power":
                        reason_core = (
                            f"Switched from {current.name} to {uni.name}; main reason: notably stronger university weight "
                            f"alongside a modest preference improvement."
                        )
                    else:
                        reason_core = (
                            f"Switched from {current.name} to {uni.name} due to a balance of higher preference and "
                            f"greater university weight."
                        )
                else:  # no clear improvement but total score higher (edge case)
                    reason_core = (
                        f"Switched from {current.name} to {uni.name} due to a small combined advantage."
                    )
                stu.match = uni.name
                current.accepted.remove(cand)
                uni.accepted.append(cand)
                matches[cand] = uni.name
                reasons[cand] = reason_core
                if current.has_free_slot():
                    free_unis.append(current)
            else:
                # Stayed with current
                pref_new_better = r_new < r_old
                power_new_better = uni_comp_new > uni_comp_old
                if pref_new_better and not power_new_better:
                    reason = (
                        f"Stayed with {current.name}; although {uni.name} proposed and is higher preference, its overall "
                        f"influence was not enough to justify a switch."
                    )
                elif power_new_better and not pref_new_better:
                    reason = (
                        f"Stayed with {current.name}; {uni.name}'s greater weight did not compensate for its lower/ equal "
                        f"preference rank."
                    )
                elif pref_new_better and power_new_better:
                    reason = (
                        f"Stayed with {current.name}; combined improvement from {uni.name} was too small to warrant change."
                    )
                else:
                    reason = (
                        f"Stayed with {current.name} because {current.name} is at least as preferred and comparably strong."
                    )
                reasons[cand] = reason

        if uni.has_free_slot() and uni.preference_pointer < len(uni.preferences_flat):
            free_unis.append(uni)

    for name, stu in students.items():
        matches[name] = stu.match
        if stu.match is None:
            reasons[name] = (
                "Unmatched; no proposing university achieved a sufficient combined advantage in preference and institutional weight."
            )

    return matches, reasons