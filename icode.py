import numpy as np
import scipy.io as sio
import scipy.signal
from PIL import Image
import numpy as ny

def photo_split(im):
    im = im[0:20, :]
    im = scipy.signal.wiener(im)
    im = im > (255 * 0.4)
    s1 = im[:, 4:17]
    s2 = im[:, 17:30]
    s3 = im[:, 30:43]
    s4 = im[:, 43:56]
    s1 = np.transpose(s1).reshape((1, 260))
    s2 = np.transpose(s2).reshape((1, 260))
    s3 = np.transpose(s3).reshape((1, 260))
    s4 = np.transpose(s4).reshape((1, 260))
    s = np.concatenate((s1, s2, s3, s4), axis=0)
    return s


def sigmoid(z):
    z = np.matrix(z)
    g = 1 / (1 + np.exp(-z))
    return g


def predict(Theta1, Theta2, Theta3, X):
    X = np.matrix(X)
    Theta1 = np.matrix(Theta1)
    Theta2 = np.matrix(Theta2)
    Theta3 = np.matrix(Theta3)
    m = X.shape[0]
    h1 = sigmoid(((np.concatenate((np.ones((m, 1)), X), axis=1)) * (np.transpose(Theta1))))
    h2 = sigmoid(((np.concatenate((np.ones((m, 1)), h1), axis=1)) * (np.transpose(Theta2))))
    h3 = sigmoid(((np.concatenate((np.ones((m, 1)), h2), axis=1)) * (np.transpose(Theta3))))
    p1 = np.ndarray.argmax(h3[0, :])
    p2 = np.ndarray.argmax(h3[1, :])
    p3 = np.ndarray.argmax(h3[2, :])
    p4 = np.ndarray.argmax(h3[3, :])
    return p1, p2, p3, p4


def openimage(fullpath):
    im = ny.array(Image.open(fullpath).convert('L'), 'f')
    return im


def loadTheta():
    matname = u'THETA.mat'
    data = sio.loadmat(matname)
    Theta1 = data['Theta1']
    Theta2 = data['Theta2']
    Theta3 = data['Theta3']
    return Theta1, Theta2, Theta3



#######################################
def verfication(fullpath):
    word = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k',
            'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    Theta1, Theta2, Theta3 = loadTheta()

    im = openimage(fullpath)
    X = photo_split(im)
    p1, p2, p3, p4 = predict(Theta1, Theta2, Theta3, X)
    verification = (word[p1[0, 0]]) + (word[p2[0, 0]]) + (word[p3[0, 0]]) + (word[p4[0, 0]])
    return verification


