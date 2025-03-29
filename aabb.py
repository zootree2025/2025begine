import random

def generate_secret_number():
    digits = list(range(10))
    random.shuffle(digits)
    return digits[:4]

def get_guess():
    while True:
        guess = input("請輸入4位數字: ")
        if len(guess) == 4 and guess.isdigit():
            return [int(d) for d in guess]
        print("輸入無效，請輸入4位數字。")

def calculate_ab(secret, guess):
    a = sum(1 for s, g in zip(secret, guess) if s == g)
    b = sum(1 for g in guess if g in secret) - a
    return a, b

def main():
    secret_number = generate_secret_number()
    attempts = 0

    print("歡迎來到幾A幾B遊戲！")
    while True:
        guess = get_guess()
        attempts += 1
        a, b = calculate_ab(secret_number, guess)
        print(f"{a}A{b}B")
        if a == 4:
            print(f"恭喜你猜對了！總共用了 {attempts} 次嘗試。")
            break

if __name__ == "__main__":
    main()