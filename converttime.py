import sys
def convertpsttoutc(time):
    #time in format HH:MM
    x = time.split(":")
    newhour = int(x[0]) + 8
    minutes = int(x[1])
    if newhour > 24:
        newhour = newhour - 24
    if newhour == 24 and minutes > 0:
        newhour = 0
    if minutes == 0 or minutes < 10:
        if newhour < 10:
            return(f"0{newhour}:0{minutes}")
        elif newhour >= 10:
            return(f"{newhour}:0{minutes}")
    else:
        if newhour < 10:
            return(f"0{newhour}:{minutes}")
        elif newhour >= 10:
            return(f"{newhour}:{minutes}")
def main():
    print(convertpsttoutc(str(sys.argv[1])))
if __name__ == '__main__':
    main()
