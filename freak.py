num_seq = [2,1,4,3]
first_score = 0
second_score = 0

for i in range(len(num_seq)):
    current_pick = num_seq.pop(0)
    if i % 2 == 0:
        first_score += current_pick
    else:
        second_score += current_pick
    if current_pick % 2 == 0:
        num_seq.reverse()

score_difference = first_score - second_score
print("The difference in scores between the first and second player is:", score_difference)


# score_difference = first_score - second_score
# print("The difference in scores between the first and second player is:", score_difference)

# def getScoreDifference(numSeq):
#     # Write your code here
#     fs = 0
#     sd = 0
#     i = 0
#     j = len(numSeq)
#     isR = False
#     turn = 0
#     while i <= j:
#         if isR:
#             if turn % 2 == 0:
#                 fs += numSeq[j]
#             else:
#                 sd += numSeq[j]
#             if numSeq[j] % 2 == 0:
#                 isR = not isR
#                 j -= 1

#         else:
#             if turn % 2 == 0:
#                 fs += numSeq[i]
#             else:
#                 sd =+numSeq[i]
#             if numSeq[i] % 2 == 0:
#                 isR = not isR
#                 i += 1
#                 turn =+1
#             return fs - sd
        
# getScoreDifference(num_seq)