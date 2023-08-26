import pyautocad
import tkinter as tk
from tkinter import messagebox
import pythoncom

class Move:
    def __init__(self, row, col):
        self.row = row
        self.col = col

computer = 'x'
player = 'o'
middle_points = [
        [25, 125], [75, 125], [125, 125], [25, 75], [75, 75], [125, 75], [25, 25], [75, 25], [125, 25] ]
text_pos = [213.3862, 64]
turn_text_pos = [213.3862, 10]


# any moves left on board?
def is_moves_left(board: list) -> bool:
    for i in range(3):
        for j in range(3):
            if board[i][j] == '-':
                return True
    return False

# check if there's any victory already
# if computer won return 10, if player won return -10, else return 0
def evaluate(board: list) -> int:
    # check rows for victory
    for row in range(3):
        if board[row][0] == board[row][1] and board[row][1] == board[row][2]:
            if board[row][0] == computer:
                return 10
            elif board[row][0] == player:
                return -10
            
    # check columns for victory.
    for col in range(3):
        if board[0][col] == board[1][col] and board[1][col] == board[2][col]:
            if board[0][col] == computer:
                return 10
            elif board[0][col] == player:
                return -10
            
    # check diagonals for victory.
    if board[0][0] == board[1][1] and board[1][1] == board[2][2]:
        if board[0][0] == computer:
            return 10
        elif board[0][0] == player:
            return -10
    if board[0][2] == board[1][1] and board[1][1] == board[2][0]:
        if board[0][2] == computer:
            return 10
        elif board[0][2] == player:
            return -10
        
    # if noone won return 0
    return 0

# minimax - consider all possible ways the game can go and return value of the board
def minimax(board: list, depth: int, is_max: bool) -> int:
    score = evaluate(board)
    # if computer or player won return the score
    if score == 10 or score == -10:
        return score
    
    # if no more moves left
    if is_moves_left(board) == False:
        return 0
    
    # if it's computer's move
    if is_max:
        best = -1000
        for i in range(3):
            for j in range(3):
                # check if spot empty
                if board[i][j] == '-':
                    board[i][j] = computer # move to check
                    best = max(best, minimax(board, depth+1, not is_max)) # recursive minimax, max val
                    board[i][j] = '-' # undo the move
        return best
    # if it's player's move
    else:
        best = 1000
        for i in range(3):
            for j in range(3):
                if board[i][j] == '-':
                    board[i][j] = player # move to check
                    best = min(best, minimax(board, depth+1, not is_max)) # recursive minimax, min val
                    board[i][j] = '-' # undo the move
        return best

# what is the best move?
def find_best_move(board: list) -> Move:
    best_val = -1000
    row = -1
    col = -1

    for i in range(3):
        for j in range(3):
            if board[i][j] == '-':
                board[i][j] = computer # move to check
                move_val = minimax(board, 0, False)
                board[i][j] = '-' # undo the move

                if move_val > best_val:
                    row = i
                    col = j
                    best_val = move_val

    return Move(row, col)

# ----------------------- end of minimax -------------------------------------

def what_field_clicked(x: float, y: float) -> int:
    field_clicked = None

    if 0 < x < 50 and 0 < y < 50:
        field_clicked = 7
    elif 50 <= x < 100 and 0 < y < 50:
        field_clicked = 8
    elif 100 <= x < 150 and 0 < y < 50:
        field_clicked = 9
    elif 0 < x < 50 and 50 <= y < 100:
        field_clicked = 4
    elif 50 <= x < 100 and 50 <= y < 100:
        field_clicked = 5
    elif 100 <= x < 150 and 50 <= y < 100:
        field_clicked = 6
    elif 0 < x < 50 and 100 <= y < 150:
        field_clicked = 1
    elif 50 <= x < 100 and 100 <= y < 150:
        field_clicked = 2
    elif 100 <= x < 150 and 100 <= y < 150:
        field_clicked = 3
    return field_clicked

# user is picking a point
def pick_point(acad: pyautocad.Autocad, board: list) -> list:
    #mark = None
    #field_clicked = None
    znakNaPlanszy = None
    kliknietePole = None
    fieldRowCol = None

    while kliknietePole==None or znakNaPlanszy==None or znakNaPlanszy=="o" or znakNaPlanszy=="x":
        point = acad.doc.Utility.GetPoint(pyautocad.APoint(0,0,0),'Wybierz pole: ')
        kliknietePole = what_field_clicked(point[0], point[1])
        if kliknietePole!=None:
            fieldRowCol = field_to_row_col_in_board(kliknietePole)
            znakNaPlanszy = board[fieldRowCol[0]][fieldRowCol[1]]
    
    miejscePunktu = middle_points[kliknietePole-1]
    board[fieldRowCol[0]][fieldRowCol[1]] = "o"
    acad.model.InsertBlock(pyautocad.APoint(miejscePunktu[0], miejscePunktu[1], 0), "kolko", 1, 1, 1, 0)
    return board

# transform field number to list row and column
def field_to_row_col_in_board(field: int) -> list:
    if field==1:
        return [0, 0]
    elif field==2:
        return [0, 1]
    elif field==3:
        return [0, 2]
    elif field==4:
        return [1, 0]
    elif field==5:
        return [1, 1]
    elif field==6:
        return [1, 2]
    elif field==7:
        return [2, 0]
    elif field==8:
        return [2, 1]
    elif field==9:
        return [2, 2]

#transform move to field number
def move_to_field(move: Move) -> int:
    if move.row==0 and move.col==0:
        return 1
    elif move.row==0 and move.col==1:
        return 2
    elif move.row==0 and move.col==2:
        return 3
    elif move.row==1 and move.col==0:
        return 4
    elif move.row==1 and move.col==1:
        return 5
    elif move.row==1 and move.col==2:
        return 6
    elif move.row==2 and move.col==0:
        return 7
    elif move.row==2 and move.col==1:
        return 8
    elif move.row==2 and move.col==2:
        return 9

if __name__ == "__main__":
    acad = pyautocad.Autocad()
    acadModel = acad.ActiveDocument.ModelSpace
    print(acad.doc.Name)

    continue_input = "Y"
    while (continue_input != "n" or continue_input != "N") and (continue_input == "Y" or continue_input == "y"):
        board = [
        ['-', '-', '-'],
        ['-', '-', '-'],
        ['-', '-', '-']
         ]
        #clear screen
        for object in acadModel:
            if object.ObjectName == "AcDbBlockReference":
                object.Delete()

        whose_turn = 9
        is_any_win = 0
        moves_left = True
        #game loop
        #if noone won yet and theres moves left
        while whose_turn!=0 and is_any_win==0 and moves_left==True:
            if whose_turn % 2 == 1: #if it's player's move
                acad.model.InsertBlock(pyautocad.APoint(turn_text_pos[0], turn_text_pos[1], 0), "turn", 1, 1, 1, 0)
                board = pick_point(acad, board)
            else: #if it's computer's move
                for object in acadModel:
                    if object.ObjectName == "AcDbBlockReference":
                        if object.Name == "turn":
                            object.Delete()
                best_move = find_best_move(board)
                board[best_move.row][best_move.col] = "x" 
                field_picked = move_to_field(best_move)
                move_coordinates = middle_points[field_picked-1]
                acad.model.InsertBlock(pyautocad.APoint(move_coordinates[0], move_coordinates[1], 0), "krzyzyk", 1, 1, 1, 0)

            is_any_win = evaluate(board)
            moves_left = is_moves_left(board)
            whose_turn = whose_turn-1

        #computer won
        if is_any_win == 10:
            acad.model.InsertBlock(pyautocad.APoint(text_pos[0], text_pos[1], 0), "lost", 1, 1, 1, 0)
        #player won
        elif is_any_win == -10:
            acad.model.InsertBlock(pyautocad.APoint(text_pos[0], text_pos[1], 0), "won", 1, 1, 1, 0)
        # draw
        elif is_any_win == 0 and moves_left == False:
            acad.model.InsertBlock(pyautocad.APoint(text_pos[0], text_pos[1], 0), "draw", 1, 1, 1, 0)
        
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        result = messagebox.askyesno("Tic-Tac-Toe", "Do you want to continue?")
        if result:
            continue_input = "Y"
        else:
            continue_input = "N"
        root.destroy()