#-*-coding=utf-8 -*-
'''
Created on 2018年6月26日
    Experiments with lab, see https://www.shiyanlou.com/courses/368/labs/1172/document
    - If run it on windows, need to download curses from https://www.lfd.uci.edu/~gohlke/pythonlibs/#curses
        and installed using command:   pip install "curses-2.2+utf8-cp36-cp36m-win_amd64.whl"
    - Cannot run within Eclipse IDE, please use cmd.exe
@author: junli
'''

from random import choice
import itertools
#import logging


import curses

# logging.basicConfig(level=logging.DEBUG,
#         format='%(asctime)s %(pathname)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s ',
#         datefmt='%a, %d %b %Y %H:%M:%S',
#         filename=r'c:\temp\test.log',
#         filemode='w')
# 
# logger = logging.getLogger()


class Action:
    UP ='up'
    DOWN = "down"
    LEFT = "left"
    RIGHT = "right"
    RESTART = "restart"
    EXIT = "exit"

    letter_codes = [ord(ch) for ch in "WSADRQwsadrq"]
    actions = [UP,DOWN,LEFT,RIGHT,RESTART,EXIT]
    directions = [UP,DOWN,LEFT,RIGHT]
    actions_dict=dict(zip(letter_codes,actions*2))
    
    def __init__(self,stdscr):
        self.stdscr = stdscr
        
    # return an action based on user input
    def get(self):
        char="I"
        while char not in self.actions_dict:
            char = self.stdscr.getch()
        return self.actions_dict[char]
    


class Grid:
    def __init__(self,size):
        self.size = size
        self.cells = None
        self.reset()
    
    # reset the grid: set all cells to 0, and randomly add two '2' cells
    def reset(self):
        self.cells = [[0 for _ in range(self.size)] for _ in range(self.size) ]
        self.add_random_item()
        self.add_random_item()
        
    # add one '2' in one of the empty cells
    def add_random_item(self):
        empty_cells = [(i,j) for i in range(self.size) for j in range(self.size) if self.cells[i][j]==0 ]
        i, j = choice(empty_cells)
        self.cells[i][j] = 2
        
    
    def invert(self):
        self.cells = [row[::-1] for row in self.cells]
        
    # B = AT,ie b(i,j) = a(j,i)
    def transpose(self):
        self.cells = [ list(row)  for row in zip(*self.cells) ]
    
    
    @staticmethod
    def tighten(row):
        new_row = [i for i in row if i>0]
        for i in range(len(row) - len(new_row)):
            new_row.append(0)
        return new_row
    
    @staticmethod
    def merge_left(row):
        score = 0
        pair = False
        new_row = []
        for i in range(len(row)):
            if pair:
                new_row.append( row[i] * 2)
                score += row[i]
#                logger.debug("merge_left, score={}".format(score))
                pair = False                
            else:            
                if i<len(row)-1 and row[i]==row[i+1]:
                    new_row.append(0)
                    pair=True
                else:                    
                    new_row.append(row[i])
        return new_row,score
    
    # return none empty elements counts
    @staticmethod    
    def non_empty_elements(row):
        return len([i for i in row if i>0])
    
    
    # (tighten , merge), (tighten, merge)... until there is no element change in row
    @staticmethod
    def move_row_left(row):       
        last_row = row
        score = 0
        while True:
            new_row,add_score = Grid.merge_left(Grid.tighten(last_row))   
#            logger.debug("move_row_left, score={},add_score={}".format(add_score,score))     
            score += add_score
            if Grid.non_empty_elements(new_row) == Grid.non_empty_elements(last_row):
                return new_row,score
            last_row = new_row
        
    
    def move_left(self):
        new_cells = []
        score = 0
        for row in self.cells:
            new_row, add_score = Grid.move_row_left(row)
            new_cells.append(new_row)
            score += add_score            
        self.cells = new_cells
        return score
        
    def move_right(self):
        self.invert()
        score = self.move_left()
        self.invert()
        return score
    
    def move_up(self):
        self.transpose()
        score =self.move_left()
        self.transpose()
        return score
        
    def move_down(self):
        self.transpose()
        score =self.move_right()
        self.transpose()
        return score
        
    @staticmethod
    def row_can_move_left(row):
        for i in range(len(row)-1):
            if row[i]==0 and row[i+1]!=0:
                return True
            if row[i]!=0 and row[i]==row[i+1]:
                return True
        return False
    
    def can_move_left(self):
        return any(Grid.row_can_move_left(row) for row in self.cells)
    
    def can_move_right(self):
        self.invert()
        can = self.can_move_left()
        self.invert()
        return can
    
    def can_move_up(self):
        self.transpose()
        can = self.can_move_left()
        self.transpose()         
        return can 
    
    def can_move_down(self):
        self.transpose()
        can = self.can_move_right()
        self.transpose()         
        return can        
    
#    def can_move(self):
#        return self.can_move_left() or self.can_move_right() or self.can_move_up() or self.can_move_down()
    def can_move(self,direction):
        return getattr(self,  'can_move_'+direction)()
    
    # actions when user gives order to move to one direction
    #   scores are stored in Grid.add_score
    # Return:  action is taken or not
    def move(self,direction):
        self.add_score = 0
        if self.can_move(direction):
            score = getattr(self,'move_'+direction)()
            self.add_random_item()
            self.add_score = score
            return True
        else:
            return False
    
    def is_over(self):
        return not any(self.can_move(direction) for direction in Action.directions)
    
    # see if any cell >= win_num
    def is_win(self,win_num):
        return  any( col>=win_num for col in itertools.chain(*self.cells))
    
    # for test purpose, print out the grid into stdout
    def _print_cells(self):
        for row in self.cells:
            print( str(row))
            
# Display grid and help information
class Screen:
    help_string1 = "(W)up (S)down (A)left (D)right"
    help_string2 = "   (R)Restart (Q)Exit"
    over_string  = "       GAME OVER"
    win_string   = "       YOU WIN!"
    
    def __init__(self, screen, grid,score, best_score,over=False,win=False):
        self.screen = screen
        self.grid = grid
        self.score = score
        self.best_score = best_score
        self.over = over
        self.win = win
    
    # display one row on scree
    def cast(self,string):
        self.screen.addstr(string + "\n")
    
    def draw_row(self,row):
        self.cast(''.join( '|{:^5}'.format(col) if col> 0 else "|     " for col in row ) + "|")
        
    def draw(self):
        self.screen.clear()
        self.cast('SCORE:{:5d}      BEST_SCORE:{:5d}' .format(self.score,self.best_score))
        for row in self.grid.cells:
            self.cast('+-----'*self.grid.size+"+")
            self.draw_row(row)
        self.cast('+-----'*self.grid.size+"+")
        
        if self.win:
            self.cast(self.win_string)
        elif self.over:
            self.cast(self.over_string)
        else:
            self.cast(self.help_string1)
        
        self.cast(self.help_string2)
        
        

class GameManager:
    
    GAME='game'
    INIT ='init'
    EXIT='exit'
    WIN='win'
    OVER='over'
    
    def __init__(self, stdscr, size=4,win_num=2048):
        self.size = size
        self.win_num = win_num
        self.action = Action(stdscr)
        self.stdscr= stdscr
        self.best_score = 0
        self.reset()
        
        
    def reset(self):
        self.state = GameManager.INIT
        self.win = False
        self.over= False
        self.score = 0
        self.grid= Grid(self.size)
        self.grid.reset()
        
    def screen(self):
        return Screen(screen=self.stdscr, score=self.score, best_score=self.best_score, grid=self.grid, win=self.win, over=self.over)
    
    
    def is_win(self):
        self.win= self.grid.is_win(self.win_num)
        return self.win

    def is_over(self):
        self.over = self.grid.is_over()
        return self.over
    
    def state_init(self):
        self.reset()
        return  GameManager.GAME
    
    def move(self,direction):
        return self.grid.move(direction)
    
    def state_game(self):
        self.screen().draw()
        action = self.action.get()
        if action == Action.RESTART:
            return GameManager.INIT
        if action == Action.EXIT:
            return GameManager.EXIT
       
        if self.move(action):
            self.score += self.grid.add_score
           
            if self.is_win():
                if self.score > self.best_score:
                    self.best_score = self.score
                return GameManager.WIN
            if self.is_over():
                if self.score > self.best_score:
                    self.best_score = self.score
                return GameManager.OVER
        return GameManager.GAME
    
    def state_win(self):
        self.screen().draw()
        action = self.action.get()
        if  action== Action.RESTART:
            return GameManager.INIT
        elif action == Action.EXIT:
            return GameManager.EXIT
        else: 
            return GameManager.WIN
        
    def state_over(self):
        self.screen().draw()
        action = self.action.get()
        if  action== Action.RESTART:
            return GameManager.INIT
        elif action == Action.EXIT:
            return GameManager.EXIT
        else: 
            return GameManager.OVER
        
    def loop(self):
        while self.state !=GameManager.EXIT:
            self.state = getattr(self, 'state_'+self.state)()
        
    
def main(stdscr):
    g = GameManager(stdscr,4,64)
    g.loop()
    
   
    
def test_2():    
    row = [2,2,4,8]
    new_row = Grid.merge_left(row)
    assert new_row == [16,0,0,0]
    
    row=[0,2,2,0]
    new_row = Grid.merge_left(row)
    assert new_row == [4,0,0,0]
    
    row=[0,2,0,0]
    new_row = Grid.merge_left(row)
    assert new_row == [2,0,0,0]
    

def test_1():
    g = Grid(4)
    g._print_cells()
    old = g.cells
    print("\n")
    g.invert()
    assert old != g.cells
    g._print_cells()
    
    g.invert()    
    assert old == g.cells
    print("\n")
    g.transpose()
    g._print_cells()
    g.transpose()
    assert old == g.cells
    
if __name__ == '__main__':    
    curses.wrapper(main )
    
    