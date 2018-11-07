import PIL.ImageGrab
import pyautogui
import win32api
import win32gui
import win32con
import time
import random


def color_hash(color):
    value = ""
    for i in range(5):
        value += "%d,%d,%d," % (color[0], color[1], color[2])
    return hash(value)


def image_hash(img):
    value = ""
    for i in range(5):
        c = img.getpixel((i * 3, i * 3))
        value += "%d,%d,%d," % (c[0], c[1], c[2])
    return hash(value)


def game_area_image_to_matrix():
    pos_to_image = {}

    for row in range(ROW_NUM):
        pos_to_image[row] = {}
        for col in range(COL_NUM):
            grid_left = col * grid_width
            grid_top = row * grid_height
            grid_right = grid_left + grid_width
            grid_bottom = grid_top + grid_height

            grid_image = game_area_image.crop((grid_left, grid_top, grid_right, grid_bottom))

            pos_to_image[row][col] = grid_image

    pos_to_type_id = {}
    image_map = {}

    empty_hash = color_hash((48, 76, 112))

    for row in range(ROW_NUM):
        pos_to_type_id[row] = {}
        for col in range(COL_NUM):
            this_image = pos_to_image[row][col]
            this_image_hash = image_hash(this_image)
            if this_image_hash == empty_hash:
                pos_to_type_id[row][col] = 0
                continue
            image_map.setdefault(this_image_hash, len(image_map) + 1)
            pos_to_type_id[row][col] = image_map.get(this_image_hash)

    return pos_to_type_id


def solve_matrix_one_step():
    for key in map:
        arr = map[key]
        arr_len = len(arr)
        for index1 in range(arr_len - 1):
            point1 = arr[index1]
            x1 = point1[0]
            y1 = point1[1]
            for index2 in range(index1 + 1, arr_len):
                point2 = arr[index2]
                x2 = point2[0]
                y2 = point2[1]
                if verifying_connectivity(x1, y1, x2, y2):
                    arr.remove(point1)
                    arr.remove(point2)
                    matrix[y1][x1] = 0
                    matrix[y2][x2] = 0
                    if arr_len == 2:
                        map.pop(key)
                    return y1, x1, y2, x2


def verifying_connectivity(x1, y1, x2, y2):
    max_y1 = y1
    while max_y1 + 1 < ROW_NUM and matrix[max_y1 + 1][x1] == 0:
        max_y1 += 1
    min_y1 = y1
    while min_y1 - 1 >= 0 and matrix[min_y1 - 1][x1] == 0:
        min_y1 -= 1

    max_y2 = y2
    while max_y2 + 1 < ROW_NUM and matrix[max_y2 + 1][x2] == 0:
        max_y2 += 1
    min_y2 = y2
    while min_y2 - 1 >= 0 and matrix[min_y2 - 1][x2] == 0:
        min_y2 -= 1

    rg_min_y = max(min_y1, min_y2)
    rg_max_y = min(max_y1, max_y2)
    if rg_max_y >= rg_min_y:
        for index_y in range(rg_min_y, rg_max_y + 1):
            min_x = min(x1, x2)
            max_x = max(x1, x2)
            flag = True
            for index_x in range(min_x + 1, max_x):
                if matrix[index_y][index_x] != 0:
                    flag = False
                    break
            if flag:
                return True

    max_x1 = x1
    while max_x1 + 1 < COL_NUM and matrix[y1][max_x1 + 1] == 0:
        max_x1 += 1
    min_x1 = x1
    while min_x1 - 1 >= 0 and matrix[y1][min_x1 - 1] == 0:
        min_x1 -= 1

    max_x2 = x2
    while max_x2 + 1 < COL_NUM and matrix[y2][max_x2 + 1] == 0:
        max_x2 += 1
    min_x2 = x2
    while min_x2 - 1 >= 0 and matrix[y2][min_x2 - 1] == 0:
        min_x2 -= 1

    rg_min_x = max(min_x1, min_x2)
    rg_max_x = min(max_x1, max_x2)
    if rg_max_x >= rg_min_x:
        for index_x in range(rg_min_x, rg_max_x + 1):
            min_y = min(y1, y2)
            max_y = max(y1, y2)
            flag = True
            for index_y in range(min_y + 1, max_y):
                if matrix[index_y][index_x] != 0:
                    flag = False
                    break
            if flag:
                return True

    return False


def execute_one_step(one_step):
    from_row, from_col, to_row, to_col = one_step

    from_x = game_area_left + (from_col + 0.5) * grid_width
    from_y = game_area_top + (from_row + 0.5) * grid_height

    to_x = game_area_left + (to_col + 0.5) * grid_width
    to_y = game_area_top + (to_row + 0.5) * grid_height

    pyautogui.moveTo(from_x, from_y)
    pyautogui.click()

    pyautogui.moveTo(to_x, to_y)
    pyautogui.click()


if __name__ == '__main__':

    COL_NUM = 19
    ROW_NUM = 11

    screen_width = win32api.GetSystemMetrics(0)
    screen_height = win32api.GetSystemMetrics(1)

    hwnd = win32gui.FindWindow(win32con.NULL, 'QQ游戏 - 连连看角色版')
    if hwnd == 0:
        exit(-1)

    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    window_left, window_top, window_right, window_bottom = win32gui.GetWindowRect(hwnd)
    if min(window_left, window_top) < 0 or window_right > screen_width or window_bottom > screen_height:
        exit(-1)
    window_width = window_right - window_left
    window_height = window_bottom - window_top

    game_area_left = window_left + 14.0 / 800.0 * window_width
    game_area_top = window_top + 181.0 / 600.0 * window_height
    game_area_right = window_left + 603 / 800.0 * window_width
    game_area_bottom = window_top + 566 / 600.0 * window_height

    game_area_width = game_area_right - game_area_left
    game_area_height = game_area_bottom - game_area_top
    grid_width = game_area_width / COL_NUM
    grid_height = game_area_height / ROW_NUM

    game_area_image = PIL.ImageGrab.grab((game_area_left, game_area_top, game_area_right, game_area_bottom))

    matrix = game_area_image_to_matrix()

    map = {}

    for y in range(ROW_NUM):
        for x in range(COL_NUM):
            grid_id = matrix[y][x]
            if grid_id == 0:
                continue
            map.setdefault(grid_id, [])
            arr = map[grid_id]
            arr.append([x, y])

    pyautogui.PAUSE = 0

    while True:
        one_step = solve_matrix_one_step()
        if not one_step:
            exit(0)
        execute_one_step(one_step)
        time.sleep(random.randint(0,0)/1000)
