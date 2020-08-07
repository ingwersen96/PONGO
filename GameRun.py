from Game import *
from NeuralNet import *

def run_game_with_ML(display, clock, weights, gen):
    
    global max_score
    
    returnScore = 0

    avg_score = 0
    test_games = 1
    score1 = 0
    steps_per_game = 250000
    score2 = 0

    if gen > 10:
        steps_per_game = steps_per_game*2
    
    for _ in range(test_games):
        
        paddle1_pos, paddle2_pos, paddle1_vel, paddle2_vel, bpos, score = init()

        count_same_direction = 0
        prev_direction = 0
        returnScore = 0
        score1_old = 0
        predicted_ball_pos = 0

        for _ in range(steps_per_game):
            
            relpos_normalized, \
            paddle_vertical_position_normalized, \
            ball_vertical_position_normalized, \
            ball_velocity = getInputs()
            
            predictions = []
            predicted_direction = np.argmax(np.array(forward_propagation(np.array([relpos_normalized,
                                                                                   paddle_vertical_position_normalized,
                                                                                   ball_vertical_position_normalized,
                                                                                   ball_velocity/40,
                                                                                   predicted_ball_pos,
                                                                                   prev_direction
                                                                                  ]
                                                                                 ).reshape(-1, 7), weights))) - 1

            if predicted_direction == prev_direction:
                count_same_direction += 1
                
            else:
                score1 = score1 + 5
                count_same_direction = 0
                prev_direction = predicted_direction
                
            if predicted_direction == 0:
                if count_same_direction >= 5:
                    score1 = score1 - 1
                
            if paddle_vertical_position_normalized * HEIGHT == 360:
                score1 = score1 - 1
            
            elif paddle_vertical_position_normalized * HEIGHT == 40:
                score1 = score1 - 1
                
            if ((ball_vertical_position_normalized - paddle_vertical_position_normalized)**2)**0.5 < 0.10:
                score1 = score1 + 2
                
            if relative_paddle_superior_position_normalized < 0.10:
                score1 = score1 + 1
            
            if relative_paddle_inferior_position_normalized < 0.10:
                score1 = score1 + 1

            new_direction = predicted_direction
            
            predicted_ctl = generate_new_direction(new_direction)
            
            returnScore_old = returnScore
            
            relpos_normalized, \
            paddle_vertical_position_normalized, \
            ball_vertical_position_normalized, \
            ball_velocity, \
            l_score, r_score, \
            returnScore, predicted_ball_pos = play_game(display, clock, gen, score, returnScore, predicted_direction, predicted_ctl)
            
            score = l_score + (score1-score1_old) + (returnScore - returnScore_old)*10000
            if score < 0:
                score = 0
                
            if score > max_score:
                max_score = score
            
            if r_score == 3:
                break
            
            score1_old = score1
                
    return score
