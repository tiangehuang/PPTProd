# encoding=utf-8
from core.PPTProd import PPTProd


if __name__ == '__main__':
    config_path = 'conf/config.yaml'
    app = PPTProd(config_path)
    app.proc()
