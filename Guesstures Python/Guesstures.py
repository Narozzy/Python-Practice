import random
from collections import deque
import os
import time
import win32com.client as wincl


""" Class Deck - Contains words for Guesstures """
class Deck:

    def __init__(self, difficulty, word_list_file=None, d1=None, d2=None):
        if word_list_file is not None:
            with open(word_list_file) as file:
                self.words = deque([line.strip() for line in file])
        else:
            """ Have passed in two Deck objects, so zip their words together """
            self.words = [list(word) for word in zip(d1.words, d2.words)]
        self.rank = difficulty
        self.dealt = deque()
    
    def shuffle(self):
        random.shuffle(self.words)
    
    def deal(self) -> [str]:
        popped = self.words.pop()
        self.dealt.appendleft(popped)
        return popped
    
    def deal_four(self):
        """ Deal four cards to the player """
        if len(self.words) > 4:
            return [self.deal() for _ in range(4)]
        else:
            self.collect
            self.shuffle
            return [self.deal() for _ in range(4)]
    
    def collect(self):
        """ Place dealt cards on top of deck in instance that deck is out """
        self.words.extend(self.dealt)
        self.dealt = []


def countdown(t):
    time.sleep(t)


def initialize():
    deck_easy = Deck("easy", "guesstures_easy_word_list.txt")
    deck_easy.shuffle()
    deck_medium = Deck("medium", "guesstures_medium_word_list.txt")
    deck_medium.shuffle()
    deck_easy_medium = Deck("Easy & Medium", d1=deck_easy, d2=deck_medium)
    return deck_easy_medium

def rules_explained(speaker):
    """ Announcement of rules """
    print("\nEach card has two options, the first is from the easier set while the second is from the harder but rewards more points.")
    speaker.Speak("Each card has two options, the first is from the easier set while the second is from the harder but rewards more points.")
    print("I will announce when the card is no longer in play, keep track of which ones you received on time.")
    speaker.Speak("I will announce when the card is no longer in play, keep track of which ones you received on time.")

def pick_difficulty_of_cards(hand):
    l = list()
    for card in hand:
        print("\n(1) Easy Set: {}".format(card[0]))
        print("(2) Hard Set: {}".format(card[1]))
        while True:
            try:
                choice = int(input("Decide between (1) or (2) then enter your choice: "))
                if (choice == 1 or choice == 2):
                    break
                else:
                    print("Input only 1 or 2.")
            except:
                print("Input only 1 or 2.")
        """ Appends easy if choice is 1 (default if incorrect input), otherwise appends harder """
        l.append(card[0]) if choice == 1 else l.append(card[1])
    return l

""" Main Function """
def main():
    deck = initialize()
    speak = wincl.Dispatch("SAPI.SpVoice")
    clear = lambda: os.system('cls')

    """ Hand Dealt """
    print("Dealing Cards")
    speak.Speak("Dealing Cards")
    hand = deck.deal_four()
    print(hand)

    """ Explanation of Rules """ 
    rules_explained(speak)

    """ Picking side to play """
    words_to_play = pick_difficulty_of_cards(hand)

    """ Game Start """
    print("Game beginning. Good Luck!")
    speak.Speak("Game beginning. Good Luck!")
    clear()
    
    for i in range(len(words_to_play)):
        print(words_to_play[i:len(words_to_play)])
        countdown(7.5)
        speak.Speak("Dropping Word {}".format(i+1))
        clear() 

if __name__ == '__main__':
    main()