import random

def generate_preferences(values):
    """
    Generate preference profiles from input values.
    values: A 2D object (like Excel sheet or pandas DataFrame) where each row
            represents an agent and columns represent alternatives.
    Returns:
        preferences: dict with agent IDs as keys and ranked alternative IDs as values.
    """
    agent_prefs = {}
    for row in range(1, values.max_row + 1):
        agent_prefs[row] = [values.cell(row, col).value for col in range(1, values.max_column + 1)]
    
    # Convert numerical values to ranks
    for agent, vals in agent_prefs.items():
        ranked = sorted(enumerate(vals, 1), key=lambda x: x[1], reverse=True)
        agent_prefs[agent] = [alt for alt, _ in ranked]
    
    return agent_prefs


def dictatorship(preferences, agent):
    """
    Dictatorship voting: Winner is the top choice of the selected agent.
    """
    if agent not in preferences:
        raise ValueError(f"Agent {agent} not found in preferences.")
    return preferences[agent][0]


def scoring_rule(preferences, score_vector, tie_break='max'):
    """
    Scoring rule voting: Assign points based on scores vector.
    """
    m = len(preferences[1])
    if len(score_vector) != m:
        raise ValueError("Score vector length does not match number of alternatives.")
    
    scores = {alt: 0 for alt in range(1, m + 1)}
    for pref in preferences.values():
        for i, alt in enumerate(pref):
            scores[alt] += score_vector[i]

    max_score = max(scores.values())
    winners = [alt for alt, val in scores.items() if val == max_score]

    return tie_break_function(preferences, tie_break, winners)


def plurality(preferences, tie_break='max'):
    """
    Plurality rule: Winner is the alternative with the most first-place votes.
    """
    top_counts = {}
    for pref in preferences.values():
        top_counts[pref[0]] = top_counts.get(pref[0], 0) + 1

    max_count = max(top_counts.values())
    winners = [alt for alt, count in top_counts.items() if count == max_count]

    return tie_break_function(preferences, tie_break, winners)


def veto(preferences, tie_break='max'):
    """
    Veto rule: Assign 0 to last place, 1 to all others.
    """
    m = len(preferences[1])
    scores = {alt: 0 for alt in preferences[1]}

    for pref in preferences.values():
        for alt in pref[:-1]:
            scores[alt] += 1

    max_score = max(scores.values())
    winners = [alt for alt, val in scores.items() if val == max_score]

    return tie_break_function(preferences, tie_break, winners)


def borda(preferences, tie_break='max'):
    """
    Borda rule: Highest points for most preferred alternative.
    """
    m = len(preferences[1])
    scores = {alt: 0 for alt in preferences[1]}

    for pref in preferences.values():
        for i, alt in enumerate(pref):
            scores[alt] += m - i

    max_score = max(scores.values())
    winners = [alt for alt, val in scores.items() if val == max_score]

    return tie_break_function(preferences, tie_break, winners)


def harmonic(preferences, tie_break='max'):
    """
    Harmonic rule: Score = 1/rank for each alternative.
    """
    scores = {}
    for pref in preferences.values():
        for i, alt in enumerate(pref):
            scores[alt] = scores.get(alt, 0) + 1 / (i + 1)

    max_score = max(scores.values())
    winners = [alt for alt, val in scores.items() if val == max_score]

    return tie_break_function(preferences, tie_break, winners)


def STV(preferences, tie_break='max'):
    """
    Single Transferable Vote: Eliminate least-preferred alternatives iteratively.
    """
    alternatives = set(preferences[1])
    while len(alternatives) > 1:
        counts = {alt: 0 for alt in alternatives}
        for pref in preferences.values():
            if pref:
                if pref[0] in counts:
                    counts[pref[0]] += 1

        min_count = min(counts.values())
        eliminated = [alt for alt, cnt in counts.items() if cnt == min_count]

        for pref in preferences.values():
            for alt in eliminated:
                if alt in pref:
                    pref.remove(alt)

        alternatives -= set(eliminated)

    return tie_break_function(preferences, tie_break, list(alternatives))


def range_voting(values, tie_break='max'):
    """
    Range voting: Select alternative with highest sum of scores.
    """
    preferences = generate_preferences(values)
    num_alts = values.max_column
    scores = {col: sum(values.cell(row, col).value for row in range(1, values.max_row + 1))
              for col in range(1, num_alts + 1)}

    max_score = max(scores.values())
    winners = [alt for alt, val in scores.items() if val == max_score]

    return tie_break_function(preferences, tie_break, winners)


def tie_break_function(preferences, tie_break, winners):
    """
    Handle tie-breaking with options: 'max', 'min', 'random', or agent-specific.
    """
    if len(winners) == 1:
        return winners[0]

    if tie_break == 'max':
        return max(winners)
    elif tie_break == 'min':
        return min(winners)
    elif tie_break == 'random':
        return random.choice(winners)
    elif tie_break in preferences:
        for alt in preferences[tie_break]:
            if alt in winners:
                return alt
    else:
        raise ValueError(f"Invalid tie_break option: {tie_break}")
