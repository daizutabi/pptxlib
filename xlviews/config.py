"""
設定ファイルを読み出す．
同じディレクトリ下にある xlviews.ini に設定パラメータを記述する．
カスタム設定は，直接 xlvies.config.rcParams を変更することで行う．
"""
import os
import configparser

config = configparser.ConfigParser()
filename = os.path.join(os.path.dirname(__file__), 'xlviews.ini')
config.read(filename)

rcParams = {}
for section in config.sections():
    for key in config[section]:
        value = config[section][key]
        if value == 'None':
            value = None
        rcParams['.'.join([section, key])] = value
