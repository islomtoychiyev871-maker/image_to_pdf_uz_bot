/*
** D05P01. Pong o'yini
** Ikki o'yinchi uchun qadam-baqadam (turn-based) ASCII Pong o'yini.
**
** Boshqaruv:
**   a / z  - chap raketkani yuqoriga / pastga surish
**   k / m  - o'ng raketkani yuqoriga / pastga surish
**   space  - navbatdagi yurishni o'tkazib yuborish
**
** Cheklovlar: system() chaqiruvlari, ko'rsatkichlar, massivlar va
** dinamik xotira ishlatilmagan. Kod struktur uslubda yozilgan.
*/

#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#define FIELD_WIDTH 80
#define FIELD_HEIGHT 25
#define PADDLE_HEIGHT 3
#define WIN_SCORE 21

#define LEFT_PADDLE_X 1
#define RIGHT_PADDLE_X (FIELD_WIDTH - 2)

#define MIN_PADDLE_Y 1
#define MAX_PADDLE_Y (FIELD_HEIGHT - 1 - PADDLE_HEIGHT)

#define TOP_BORDER_Y 0
#define BOTTOM_BORDER_Y (FIELD_HEIGHT - 1)
#define LEFT_BORDER_X 0
#define RIGHT_BORDER_X (FIELD_WIDTH - 1)

#define BALL_START_X (FIELD_WIDTH / 2)
#define BALL_START_Y (FIELD_HEIGHT / 2)
#define PADDLE_START_Y ((FIELD_HEIGHT - PADDLE_HEIGHT) / 2)

typedef struct
{
    int ball_x;
    int ball_y;
    int ball_dx;
    int ball_dy;
    int left_paddle_y;
    int right_paddle_y;
    int left_score;
    int right_score;
} GameState;

static int random_horizontal_direction(void)
{
    if (rand() % 2 == 0)
    {
        return 1;
    }
    return -1;
}

static int random_vertical_direction(void)
{
    return (rand() % 3) - 1;
}

static GameState reset_ball(GameState state)
{
    state.ball_x = BALL_START_X;
    state.ball_y = BALL_START_Y;
    state.ball_dx = random_horizontal_direction();
    state.ball_dy = random_vertical_direction();
    return state;
}

static GameState init_game(void)
{
    GameState state;

    srand((unsigned int)time(NULL));

    state.left_paddle_y = PADDLE_START_Y;
    state.right_paddle_y = PADDLE_START_Y;
    state.left_score = 0;
    state.right_score = 0;
    state = reset_ball(state);

    return state;
}

static int is_paddle_hit(int paddle_y, int ball_y)
{
    if (ball_y >= paddle_y && ball_y <= paddle_y + PADDLE_HEIGHT - 1)
    {
        return 1;
    }
    return 0;
}

static int clamp_paddle_y(int paddle_y)
{
    if (paddle_y < MIN_PADDLE_Y)
    {
        return MIN_PADDLE_Y;
    }
    if (paddle_y > MAX_PADDLE_Y)
    {
        return MAX_PADDLE_Y;
    }
    return paddle_y;
}

static GameState process_input(GameState state, int input_char)
{
    if (input_char == 'a' || input_char == 'A')
    {
        state.left_paddle_y = clamp_paddle_y(state.left_paddle_y - 1);
    }
    else if (input_char == 'z' || input_char == 'Z')
    {
        state.left_paddle_y = clamp_paddle_y(state.left_paddle_y + 1);
    }
    else if (input_char == 'k' || input_char == 'K')
    {
        state.right_paddle_y = clamp_paddle_y(state.right_paddle_y - 1);
    }
    else if (input_char == 'm' || input_char == 'M')
    {
        state.right_paddle_y = clamp_paddle_y(state.right_paddle_y + 1);
    }

    return state;
}

static GameState bounce_off_top_and_bottom(GameState state)
{
    int next_y;

    next_y = state.ball_y + state.ball_dy;
    if (next_y <= TOP_BORDER_Y || next_y >= BOTTOM_BORDER_Y)
    {
        state.ball_dy = -state.ball_dy;
    }

    return state;
}

static GameState move_ball(GameState state)
{
    int next_x;
    int next_y;

    state = bounce_off_top_and_bottom(state);

    next_x = state.ball_x + state.ball_dx;
    next_y = state.ball_y + state.ball_dy;

    if (next_x <= LEFT_PADDLE_X)
    {
        if (is_paddle_hit(state.left_paddle_y, next_y))
        {
            state.ball_dx = 1;
            state.ball_dy = random_vertical_direction();
            state.ball_y = next_y;
            state.ball_x = LEFT_PADDLE_X + 1;
        }
        else
        {
            state.right_score = state.right_score + 1;
            state = reset_ball(state);
        }
    }
    else if (next_x >= RIGHT_PADDLE_X)
    {
        if (is_paddle_hit(state.right_paddle_y, next_y))
        {
            state.ball_dx = -1;
            state.ball_dy = random_vertical_direction();
            state.ball_y = next_y;
            state.ball_x = RIGHT_PADDLE_X - 1;
        }
        else
        {
            state.left_score = state.left_score + 1;
            state = reset_ball(state);
        }
    }
    else
    {
        state.ball_x = next_x;
        state.ball_y = next_y;
    }

    return state;
}

static char get_cell_char(GameState state, int row, int col)
{
    if (row == TOP_BORDER_Y || row == BOTTOM_BORDER_Y)
    {
        return '-';
    }
    if (col == LEFT_BORDER_X || col == RIGHT_BORDER_X)
    {
        return '|';
    }
    if (col == LEFT_PADDLE_X && is_paddle_hit(state.left_paddle_y, row))
    {
        return '#';
    }
    if (col == RIGHT_PADDLE_X && is_paddle_hit(state.right_paddle_y, row))
    {
        return '#';
    }
    if (row == state.ball_y && col == state.ball_x)
    {
        return 'O';
    }
    return ' ';
}

static void draw_field(GameState state)
{
    int row;
    int col;

    printf("\033[2J\033[H");

    for (row = 0; row < FIELD_HEIGHT; row++)
    {
        for (col = 0; col < FIELD_WIDTH; col++)
        {
            putchar(get_cell_char(state, row, col));
        }
        putchar('\n');
    }

    printf("Hisob: %d : %d\n", state.left_score, state.right_score);
    printf("Boshqaruv: a/z - chap raketka, k/m - o'ng raketka, ");
    printf("space - yurishni o'tkazib yuborish\n");
    printf("Yurish kiriting: ");
    fflush(stdout);
}

static int is_game_over(GameState state)
{
    if (state.left_score >= WIN_SCORE || state.right_score >= WIN_SCORE)
    {
        return 1;
    }
    return 0;
}

static void print_winner(GameState state)
{
    printf("\nO'yin tugadi! Yakuniy hisob: %d : %d\n",
           state.left_score, state.right_score);

    if (state.left_score > state.right_score)
    {
        printf("Chap o'yinchi g'olib bo'ldi!\n");
    }
    else if (state.right_score > state.left_score)
    {
        printf("O'ng o'yinchi g'olib bo'ldi!\n");
    }
    else
    {
        printf("O'yin durrang bilan tugadi.\n");
    }
}

static void flush_input_line(void)
{
    int ch;

    ch = getchar();
    while (ch != '\n' && ch != EOF)
    {
        ch = getchar();
    }
}

static int read_input(void)
{
    int ch;

    ch = getchar();
    if (ch != '\n' && ch != EOF)
    {
        flush_input_line();
    }

    return ch;
}

int main(void)
{
    GameState state;
    int input_char;

    state = init_game();

    while (!is_game_over(state))
    {
        draw_field(state);

        input_char = read_input();
        if (input_char == EOF)
        {
            break;
        }

        state = process_input(state, input_char);
        state = move_ball(state);
    }

    draw_field(state);
    print_winner(state);

    return 0;
}
