import openpyxl
import matplotlib.pyplot as plt

# Excel file load
wb = openpyxl.load_workbook("emotion_log.xlsx")
ws = wb.active

# Emotion count dictionary
emotion_count = {
    "happy": 0,
    "sad": 0,
    "angry": 0,
    "neutral": 0,
    "surprise": 0,
    "fear": 0,
    "disgust": 0
}

# Read Excel data
for row in ws.iter_rows(min_row=2, values_only=True):
    emotion = row[2]
    if emotion in emotion_count:
        emotion_count[emotion] += 1

# Remove zero values
emotion_count = {k: v for k, v in emotion_count.items() if v > 0}

# Data
labels = list(emotion_count.keys())
sizes = list(emotion_count.values())

# Colors
colors = [
    "green",
    "blue",
    "red",
    "gray",
    "orange",
    "purple",
    "brown"
]

# Bar Graph
plt.figure(figsize=(10, 5))

plt.subplot(1, 2, 1)
plt.bar(labels, sizes, color=colors[:len(labels)])
plt.title("Emotion Bar Graph")
plt.xlabel("Emotion")
plt.ylabel("Count")

# Pie Chart
plt.subplot(1, 2, 2)
plt.pie(
    sizes,
    labels=labels,
    autopct="%1.1f%%",
    colors=colors[:len(labels)]
)
plt.title("Emotion Pie Chart")

# Save image
plt.savefig("graph.png")

# Show graph
plt.show()