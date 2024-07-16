import matplotlib.pyplot as plt
import numpy


def get_val(i):                         #this function returns a list containg the student's marks in all tests
    val1 = [i.get('cie1'), i.get('cie2'), i.get('cie3'), i.get('aat1'), i.get('aat2'), i.get('cie'), i.get('aat'), i.get('internals')]
    return val1


def chart(att, val1, val2, val3, val4, name):                      #this function draws the student's performance analysis chart(PAC)
    plt.cla()                                                      #clears the axes of the PAC
    plt.clf()                                                      #clears the plot of PAC
    width = 0.15                                                   #setting the width of each bar in the PAC
    values = numpy.arange(len(att))                                #to obtain a range of values to position the bars in the PAC

    plt.bar(values, val1, width, label="Student's Marks")               #plotting individual bars of the PAC
    plt.bar(values+width, val2, width, label="Class Average Marks")
    plt.bar(values+(2*width), val3, width, label="Class Highest Marks")
    plt.bar(values+(3*width), val4, width, label="Maximum Marks")

    plt.xlabel('Categories of Marks')                              #setting the x-axis label for the PAC
    plt.ylabel('Marks')                                            #setting the y-axis label for the PAC
    plt.title(f"{name}'s Internals Performance Analysis")          #setting the title of the PAC
    plt.legend()                                                   #setting the legend for the PAC
    plt.xticks(values+0.225, att)                                  #labelling each setion of bars as test headings

    plt.draw()                                                     #plotting the PAC
    plt.savefig(f"{name}.png")                                     #saving the PAC into the project folder to attach it with the mail
