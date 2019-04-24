"""
Decision tree example using scikit-learn
"""
from sklearn.tree import DecisionTreeClassifier
from sklearn.model_selection import train_test_split
import pickle
import pandas as pd
import os


_zoo_classifications = {
    1: "mammal",
    2: "bird",
    3: "reptile",
    4: "fish",
    5: "amphibian",
    6: "insect",
    7: "mollusc"
}


#if __name__ == "__main__":

# Load the zoo data
dataset = pd.read_csv(os.path.join(os.path.dirname(__file__), "..", "data", "zoo.csv"))

# Drop the animal names since this is not a good feature to split the data on
dataset = dataset.drop("animal_name", axis=1)

# Split the data into a training and a testing set
features = dataset.drop("class", axis=1)
targets = dataset["class"]

train_features, test_features, train_targets, test_targets = \
    train_test_split(features, targets, train_size=0.75, random_state=245245)

# Train the model
tree = DecisionTreeClassifier(criterion="entropy", max_depth=5)
tree = tree.fit(train_features, train_targets)

ptree = pickle.dumps(tree)
tree = pickle.loads(ptree)

# Predict the classes of new, unseen data
prediction = tree.predict(test_features)

# Check the accuracy
scrore = tree.score(test_features, test_targets) * 100
print("The prediction accuracy is: {:0.2f}%".format(scrore))

# Try predicting based on some features
features = {
    "hair": 0,
    "feathers": 1,
    "eggs": 1,
    "milk": 0,
    "airbone": 1,
    "aquatic": 0,
    "predator": 0,
    "toothed": 1,
    "backbone": 1,
    "breathes": 1,
    "venomous": 0,
    "fins": 0,
    "legs": 1,
    "tail": 1,
    "domestic": 0,
    "catsize": 0
}

features = pd.DataFrame([features], columns=train_features.columns)
prediction = tree.predict(features)[0]
print("Best guess is {}".format(_zoo_classifications[prediction]))

# Get the probabilities for each class
proba = tree.predict_proba(features)[0]
index = [_zoo_classifications[c] for c in tree.classes_]
print(pd.Series(proba, index=index))

