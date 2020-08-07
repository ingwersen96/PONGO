#PONG pygame
import random
import pygame, sys
from pygame.locals import *
import time

#colors
WHITE = (255,255,255)
RED = (255,0,0)
GREEN = (0,255,0)
BLACK = (0,0,0)

#globals
WIDTH = 600
HEIGHT = 400       
BALL_RADIUS = 20
PAD_WIDTH = 8
PAD_HEIGHT = 80
HALF_PAD_WIDTH = PAD_WIDTH // 2
HALF_PAD_HEIGHT = PAD_HEIGHT // 2
ball_pos = [0,0]
ball_vel = [0,0]
paddle1_vel = 0
paddle2_vel = 0
l_score = 0
r_score = 0
nScore = 0
max_score = 0
gen = 1
returnScore = 0

max_scr = 0

pygame.init()

#canvas declaration
window = pygame.display.set_mode((WIDTH, HEIGHT), 0, 32)
pygame.display.set_caption('Hello World')
display=window
clock=pygame.time.Clock()

# helper function that spawns a ball, returns a position vector and a velocity vector
# if right is True, spawn to the right, else spawn to the left
def ball_init(right):
    global ball_pos, ball_vel # these are vectors stored as lists
    ball_pos = [WIDTH//2,HEIGHT//2]
    horz = random.randrange(20,40)
    vert = random.randrange(10,20)
    
    if right == False:
        horz = - horz
        
    ball_vel = [horz,-vert]
    return ball_pos

# define event handlers
def init():
    global paddle1_pos, paddle2_pos, paddle1_vel, paddle2_vel,l_score,r_score  # these are floats
    global score1, score2  # these are ints
    paddle1_pos = [HALF_PAD_WIDTH - 1,HEIGHT//2]
    paddle2_pos = [WIDTH +1 - HALF_PAD_WIDTH,HEIGHT//2]
    l_score = 0
    r_score = 0
    if random.randrange(0,2) == 0:
        bpos = ball_init(True)
    else:
        bpos = ball_init(False)
        
    return paddle1_pos, paddle2_pos, paddle1_vel, paddle2_vel, bpos, l_score


#draw function of canvas
def draw(canvas, returnScore, gen, score, predicted_dir, predicted_ctl):
    global paddle1_pos, paddle2_pos, ball_pos, ball_vel, l_score, r_score
           
    canvas.fill(BLACK)
    pygame.draw.line(canvas, WHITE, [WIDTH // 2, 0],[WIDTH // 2, HEIGHT], 1)
    pygame.draw.line(canvas, WHITE, [PAD_WIDTH, 0],[PAD_WIDTH, HEIGHT], 1)
    pygame.draw.line(canvas, WHITE, [WIDTH - PAD_WIDTH, 0],[WIDTH - PAD_WIDTH, HEIGHT], 1)
    pygame.draw.circle(canvas, WHITE, [WIDTH//2, HEIGHT//2], 70, 1)

    # update paddle's vertical position, keep paddle on the screen
    if paddle1_pos[1] > HALF_PAD_HEIGHT and paddle1_pos[1] < HEIGHT - HALF_PAD_HEIGHT:
        paddle1_pos[1] += paddle1_vel
    elif paddle1_pos[1] == HALF_PAD_HEIGHT and paddle1_vel > 0:
        paddle1_pos[1] += paddle1_vel
    elif paddle1_pos[1] == HEIGHT - HALF_PAD_HEIGHT and paddle1_vel < 0:
        paddle1_pos[1] += paddle1_vel
    
    if paddle2_pos[1] > HALF_PAD_HEIGHT and paddle2_pos[1] < HEIGHT - HALF_PAD_HEIGHT:
        paddle2_pos[1] += paddle2_vel
    elif paddle2_pos[1] == HALF_PAD_HEIGHT and paddle2_vel > 0:
        paddle2_pos[1] += paddle2_vel
    elif paddle2_pos[1] == HEIGHT - HALF_PAD_HEIGHT and paddle2_vel < 0:
        paddle2_pos[1] += paddle2_vel

    #update ball
    ball_pos[0] += int(ball_vel[0])
    ball_pos[1] += int(ball_vel[1])

    #draw paddles and ball
    pygame.draw.circle(canvas, RED, ball_pos, 20, 0)
    pygame.draw.polygon(canvas, GREEN, [[paddle1_pos[0] - HALF_PAD_WIDTH, paddle1_pos[1] - HALF_PAD_HEIGHT], [paddle1_pos[0] - HALF_PAD_WIDTH, paddle1_pos[1] + HALF_PAD_HEIGHT], [paddle1_pos[0] + HALF_PAD_WIDTH, paddle1_pos[1] + HALF_PAD_HEIGHT], [paddle1_pos[0] + HALF_PAD_WIDTH, paddle1_pos[1] - HALF_PAD_HEIGHT]], 0)
    pygame.draw.polygon(canvas, GREEN, [[paddle2_pos[0] - HALF_PAD_WIDTH, paddle2_pos[1] - HALF_PAD_HEIGHT], [paddle2_pos[0] - HALF_PAD_WIDTH, paddle2_pos[1] + HALF_PAD_HEIGHT], [paddle2_pos[0] + HALF_PAD_WIDTH, paddle2_pos[1] + HALF_PAD_HEIGHT], [paddle2_pos[0] + HALF_PAD_WIDTH, paddle2_pos[1] - HALF_PAD_HEIGHT]], 0)

    #ball collision check on top and bottom walls
    if int(ball_pos[1]) <= BALL_RADIUS:
        ball_vel[1] = - ball_vel[1]
    if int(ball_pos[1]) >= HEIGHT + 1 - BALL_RADIUS:
        ball_vel[1] = -ball_vel[1]
    
    #ball collison check on gutters or paddles
    if int(ball_pos[0]) <= BALL_RADIUS + PAD_WIDTH and int(ball_pos[1]) in range(paddle1_pos[1] - HALF_PAD_HEIGHT,paddle1_pos[1] + HALF_PAD_HEIGHT,1):
        ball_vel[0] = -ball_vel[0]
        ball_vel[0] *= 1.1
        ball_vel[1] *= 1.1
        returnScore += 1
    elif int(ball_pos[0]) <= BALL_RADIUS + PAD_WIDTH:
        r_score += 1
        ball_init(True)
        
    if int(ball_pos[0]) >= WIDTH + 1 - BALL_RADIUS - PAD_WIDTH and int(ball_pos[1]) in range(paddle2_pos[1] - HALF_PAD_HEIGHT,paddle2_pos[1] + HALF_PAD_HEIGHT,1):
        ball_vel[0] = -ball_vel[0]
        ball_vel[0] *= 1.1
        ball_vel[1] *= 1.1
    elif int(ball_pos[0]) >= WIDTH + 1 - BALL_RADIUS - PAD_WIDTH:
        l_score += 1
        ball_init(False)
        
    predicted_key = ""
    if predicted_ctl == 40:
        predicted_key = "Down"
    elif predicted_ctl == -40:
        predicted_key = "Up"
    elif predicted_ctl == 0:
        predicted_key = "Stop"

    #update scores
    myfont1 = pygame.font.SysFont("Comic Sans MS", 20)
    label1 = myfont1.render("Score "+str(l_score), 1, (255,255,0))
    canvas.blit(label1, (50,20))

    myfont2 = pygame.font.SysFont("Comic Sans MS", 20)
    label2 = myfont2.render("Score "+str(r_score), 1, (255,255,0))
    canvas.blit(label2, (470, 20))  
    
    
    d_up, d_down = relativePos(paddle1_pos, ball_pos)
    
    myfont4 = pygame.font.SysFont("Comic Sans MS", 15)
    label4 = myfont4.render("Rel. Pos. Up: "+str(d_up), 1, (255,255,0))
    canvas.blit(label4, (50, 360))  
    
    myfont5 = pygame.font.SysFont("Comic Sans MS", 15)
    label5 = myfont5.render("Rel. Pos. Down: "+str(d_down), 1, (255,255,0))
    canvas.blit(label5, (50, 380))  
    
    myfont6 = pygame.font.SysFont("Comic Sans MS", 15)
    label6 = myfont6.render("Paddle Y Pos.: "+str(paddle1_pos[1]), 1, (255,255,0))
    canvas.blit(label6, (50, 340))  
    
    myfont7 = pygame.font.SysFont("Comic Sans MS", 15)
    label7 = myfont7.render("Ball Y Pos.: "+str(ball_pos[1]), 1, (255,255,0))
    canvas.blit(label7, (50, 320))  
    
    myfont9 = pygame.font.SysFont("Comic Sans MS", 15)
    label9 = myfont9.render("Generation.: "+str(gen), 1, (255,255,0))
    canvas.blit(label9, (280, 40))    
    
    myfont3 = pygame.font.SysFont("Comic Sans MS", 20)
    label3 = myfont3.render("Num of Returns: " + str(returnScore), 1, (255,255,0))
    canvas.blit(label3, (280, 20))  
    
    myfont10 = pygame.font.SysFont("Comic Sans MS", 15)
    label10 = myfont10.render("NN Score: "+str(score), 1, (255,255,0))
    canvas.blit(label10, (470, 380))  
    
    myfont11 = pygame.font.SysFont("Comic Sans MS", 15)
    label11 = myfont11.render("NN Answer: "+str(predicted_dir), 1, (255,255,0))
    canvas.blit(label11, (470, 360))  
    
    myfont12 = pygame.font.SysFont("Comic Sans MS", 15)
    label12 = myfont12.render("Predicted: "+str(predicted_key), 1, (255,255,0))
    canvas.blit(label12, (470, 340))
    
    
    new_vel = 0
    #ball collision check on top and bottom walls
    if int(ball_pos[1]) <= BALL_RADIUS:
        new_vel = - ball_vel[1]
    if int(ball_pos[1]) >= HEIGHT + 1 - BALL_RADIUS:
        new_vel = -ball_vel[1]
    
    predicted_ball_pos = ball_pos[1]
    predicted_ball_pos += int(new_vel)
    
    #ball collision check on top and bottom walls
    if int(predicted_ball_pos) <= BALL_RADIUS:
        new_vel = - new_vel
    if int(predicted_ball_pos) >= HEIGHT + 1 - BALL_RADIUS:
        new_vel = - new_vel
        
    predicted_ball_pos += int(new_vel)
    
    predicted_ball_pos = predicted_ball_pos // HEIGHT
        
    return returnScore, predicted_ball_pos
    
#keydown handler
def keydown(event):
    global paddle1_vel, paddle2_vel
    
    if event.key == K_UP:
        paddle2_vel = -8
    elif event.key == K_DOWN:
        paddle2_vel = 8
    elif event.key == K_w:
        paddle1_vel = -8
    elif event.key == K_s:
        paddle1_vel = 8
        
#keyup handler
def keyup(event):
    global paddle1_vel, paddle2_vel
    
    if event.key in (K_w, K_s):
        paddle1_vel = 0
    elif event.key in (K_UP, K_DOWN):
        paddle2_vel = 0


def relativePos(paddle_pos, ball_pos):
    
    relpos = ((ball_pos[1] - paddle_pos[1])**2 + (ball_pos[0] - paddle_pos[0])**2)**0.5
    
    
    MAX = (WIDTH**2 + HEIGHT**2)**0.5
    
    
    relpos_normalized = relpos / MAX
    
    return relpos_normalized

def getInputs():
    
    global paddle1_pos, ball_pos, ball_vel
    relpos_normalized = relativePos(paddle1_pos, ball_pos)
    
    b_vel = (ball_vel[0]**2 + ball_vel[1]**2)**0.5
    
    return relpos_normalized, paddle1_pos[1]/HEIGHT, ball_pos[1]/HEIGHT, b_vel

def controls(key_val):
    
    global paddle1_vel
    
    if key_val == 0:
        paddle1_vel = 8
        
    elif key_val == 1:
        paddle1_vel = -8
        
    elif key_val == 2:
        paddle1_vel = 0

def nnScore(nScore):
    
    nScore = nScore + 1
    
def generate_new_direction(new_direction):
    global paddle1_vel
    
    if new_direction == -1:
        paddle1_vel = -40
    
    elif new_direction == 0:
        paddle1_vel = 0
    
    elif new_direction == 1:
        paddle1_vel = 40
        
    return paddle1_vel
    
def play_game(display, clock, gen, last_score, returnScore, predicted_dir, predicted_ctl):
    
    start_time = time.time()
    
    while r_score <= 3:

        score = l_score*100 + ((time.time() - start_time)) + last_score
        
        returnScore, predicted_ball_pos = draw(window, returnScore, gen, score, predicted_dir, predicted_ctl)

        dup, ddown, p1y, by,bv = getInputs()
    
        paddle2_pos[1] = ball_pos[1]
        
        relpos_normalized, \
        paddle_vertical_position_normalized, \
        ball_vertical_position_normalized, \
        ball_velocity = getInputs()

        for event in pygame.event.get():
            if event.type == KEYDOWN:
                keydown(event)
            elif event.type == KEYUP:
                keyup(event)
            elif event.type == QUIT:
                pygame.quit()
                sys.exit()            

        pygame.display.update()
        clock.tick(50000)
        pygame.init()
        clock=pygame.time.Clock()
        
        return relative_paddle_superior_position_normalized, \
    relative_paddle_inferior_position_normalized, \
    paddle_vertical_position_normalized, \
    ball_vertical_position_normalized, \
    ball_velocity, \
    score, \
    r_score, \
    returnScore, predicted_ball_pos
