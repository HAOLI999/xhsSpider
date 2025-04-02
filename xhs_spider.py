from DrissionPage import ChromiumPage
from urllib.parse import quote
import pandas as pd
import time
import random
from tqdm import tqdm

def sign_in():
    """登录小红书"""
    page = ChromiumPage()
    page.get('https://www.xiaohongshu.com')
    print("请扫码登录")
    time.sleep(20)

def search(keyword):
    """搜索指定关键词"""
    global page
    page = ChromiumPage()
    page.get(f'https://www.xiaohongshu.com/search_result?keyword={keyword}&source=web_search_result_notes')

def page_scroll_down():
    """向下滑动页面以加载更多数据"""
    print("********下滑页面********")
    random_time = random.uniform(0.5, 1.5)
    time.sleep(random_time)
    page.scroll.to_bottom()

def view_posts(scroll_times):
    """浏览所有收集到的帖子"""
    print(f"开始浏览，下滑次数: {scroll_times}...")
    
    # 存储所有帖子链接
    all_post_links = []
    
    # 先下滑加载所有帖子
    for scroll_count in range(scroll_times):
        page_scroll_down()
        time.sleep(3)
        
        # 收集当前可见的帖子链接
        current_post_links = []
        try:
            note_links = page.eles('css:a.cover.mask.ld')
            for note_link in note_links:
                try:
                    href = note_link.attr('href')
                    if href:
                        full_url = f'https://www.xiaohongshu.com{href}' if not href.startswith('http') else href
                        if full_url not in all_post_links:  # 避免重复
                            current_post_links.append(full_url)
                except:
                    continue
            
            # 添加到总链接列表
            all_post_links.extend(current_post_links)
            print(f"第 {scroll_count+1} 次下滑后，共收集到 {len(all_post_links)} 篇帖子链接")
            
        except Exception as e:
            print(f"收集帖子链接时出错: {e}")
            continue
    
    # 修改这行代码，将post_links改为all_post_links
    print(f"共收集到 {len(all_post_links)} 篇帖子链接")
    
    # 初始化数据存储
    all_comments = []
    
    # 依次浏览每个帖子
    # 只需将 post_links 替换为 all_post_links
    for i, url in enumerate(all_post_links, 1):
        try:
            # 在当前标签页打开帖子
            page.get(url)
            print(f"正在浏览第 {i} 篇帖子...")
            
            # 获取评论内容
            comments_data = []
            try:
                # 使用XPath定位评论容器
                comments_container = page.ele('xpath:/html/body/div[2]/div[1]/div[2]/div[2]/div/div[1]/div[4]/div[2]/div[3]/div', timeout=10)
                if not comments_container:
                    print("未找到评论容器")
                    return all_comments
                    
                # 获取所有评论项
                comment_items = comments_container.eles('xpath:.//div[contains(@id, "comment-")]')
                print(f"找到 {len(comment_items)} 条评论")
                
                for item in comment_items:
                    try:
                        # 更健壮的评论内容获取方式
                        content_ele = item.ele('xpath:.//div[contains(@class,"content")]//span[contains(@class,"text")]', timeout=2)
                        content = content_ele.text if content_ele else "无内容"
                        
                        # 检查评论内容是否包含"长沙"
                        if "长沙" not in content:
                            continue
                        
                        # 更健壮的用户名获取方式
                        username_ele = item.ele('xpath:.//div[contains(@class,"author")]//a[contains(@class,"name")]', timeout=2)
                        username = username_ele.text if username_ele else "匿名用户"
                        
                        print(f"获取到评论 - 用户: {username}, 内容: {content}")
                        
                        comments_data.append({
                            'username': username,
                            'content': content,
                            'like_count': item.ele('xpath:.//div[contains(@class,"interactions")]//span[contains(@class,"count")]').text if item.ele('xpath:.//div[contains(@class,"interactions")]//span[contains(@class,"count")]', timeout=0.5) else '0',
                            'date': item.ele('xpath:.//div[contains(@class,"date")]/span').text if item.ele('xpath:.//div[contains(@class,"date")]/span', timeout=0.5) else ''
                        })
                    except Exception as e:
                        print(f"解析评论时出错: {e}")
                        continue
                    
            except Exception as e:
                print(f"获取评论容器时出错: {e}")
            
            # 保存评论数据
            all_comments.append({
                'note_url': url,
                'comments': comments_data
            })
            
            # 返回搜索结果页
            # page.back()
            # print(f"已返回搜索结果页，等待3秒...")
            # time.sleep(3)
            
        except Exception as e:
            print(f"浏览帖子时出错: {e}")
            continue
            
    print(f"已获取 {len(all_comments)} 篇笔记的评论")
    return all_comments

def craw(scroll_times):  # 参数改为下滑次数
    """循环下滑页面并浏览帖子"""
    all_data = []
    comments_data = view_posts(scroll_times)  # 只需调用一次
    if comments_data:
        all_data.extend(comments_data)
    
    # 调试输出
    print(f"\n最终收集的数据结构: {all_data}")
    print(f"总笔记数: {len(all_data)}")
    print(f"总评论数: {sum(len(item['comments']) for item in all_data if 'comments' in item)}")
    
    return all_data

def save_to_excel(data, filename="comments_results.xlsx"):
    """保存评论数据到 Excel 文件"""
    # 展开评论数据
    expanded_data = []
    for item in data:
        note_url = item['note_url']
        for comment in item['comments']:
            expanded_data.append({
                '笔记链接': note_url,
                '用户名': comment.get('username', ''),
                '评论内容': comment.get('content', ''),
                '点赞数': comment.get('like_count', ''),
                '发布时间': comment.get('date', '')
            })
    
    df = pd.DataFrame(expanded_data)
    df.to_excel(filename, index=False)
    print(f"共保存了 {len(df)} 条评论数据到 {filename}")

if __name__ == '__main__':
    # 搜索关键词
    keyword = "宠物搭车"
    
    # 自定义浏览参数
    scroll_times = 5  # 下滑次数
    
    # 转为 URL 编码
    keyword_encode = quote(keyword)

    # 搜索笔记
    search(keyword_encode)

    # 开始浏览并获取数据
    comments_data = craw(scroll_times)  # 传入下滑次数
    
    # 调试输出返回的数据
    print(f"\n返回的comments_data数据结构: {comments_data}")
    print(f"返回数据中的笔记数: {len(comments_data) if comments_data else 0}")
    
    # 保存评论到Excel
    if comments_data and any(len(item.get('comments', [])) > 0 for item in comments_data):  # 更安全的检查
        save_to_excel(comments_data)
    else:
        print("没有获取到任何评论数据")