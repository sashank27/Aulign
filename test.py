import sys
import os
import numpy as np
import pickle
import argparse


def extractZernikeCoefficents(file):
        import codecs
        import re

        pattern = "^Z\s*(\d*)\s*([-]?[0-9]+[,.]?[0-9]*).*$"

        x = []
        for line in codecs.open(file, 'r', encoding='utf16'):
            match = re.search(pattern, line)
            if match is not None:
                x.append(float(match.group(2)))

        coefficients = np.array(x)
        return coefficients


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-f", "--file", help="path of the file contaning the zernike coefficient (from Zemax configuration)")
    args = parser.parse_args()

    filename = args.file
    coefficients = extractZernikeCoefficents(filename)

    X = coefficients[:16]
    model_class = pickle.load(open("rfc_model.pkl", "rb"))

    reg0 = pickle.load(open("reg_pm.pkl", "rb"))
    reg1 = pickle.load(open("reg_sm.pkl", "rb"))
    reg2 = pickle.load(open("reg_lens.pkl", "rb"))
    reg3 = pickle.load(open("reg_ccd.pkl", "rb"))

    print('Proabability of misalignment of each component (with specific parameter in each component):-')
    prob = model_class.predict_proba(X.reshape(1, -1))[0] * 100

    elements = ['Primary Mirror', 'Secondary Mirror', 'Lens', 'CCD']
    regmodels = [reg0, reg1, reg2, reg3]

    print()
    for i in range(4):
        print("%s -> %0.2f %%" % (elements[i], prob[i]))
        comp_prob = regmodels[i].predict(X.reshape(1, -1))[0]

        print('Decenter in X:', comp_prob[0])
        print('Decenter in Y:', comp_prob[1])
        print('Tilt about X:', comp_prob[2])
        print('Tilt about Y:', comp_prob[3])
        print('-----------------------')
        print()
