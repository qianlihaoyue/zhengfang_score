<h2>Installation</h2>

`git clone https://github.com/qianlihaoyue/zhengfang_score.git`

`pip install -r requriements.txt`

<h2>Normal</h2>

`python main.py`

<h2>No input</h2>

Modify `main.py` line 144 

```python
if __name__ == "__main__":
    url = 'http://202.206.243.3'
    user=input('学号：')
    print('密码不可见')
    pswd=getpass.getpass('password:')
    #user = '2018xxxxxxxx'
    #pswd = "xxxxxxxxxxxx"
```

<h2>Note</h2>

1.如果登陆不成功，可能是密码输错或者验证码识别错误，多试一次就行。

2.仅实用于燕山大学教务系统，禁止攻击教务系统！！！

3.参考 <a href="https://github.com/ZYSzys/ZhengFang_System_Spider">ZhengFang_System_Spider</a>  and <a href="https://github.com/mepeichun/check_score_system">check_score_system</a>，我只是搬砖罢了
