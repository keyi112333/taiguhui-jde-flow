
import functools
import time

from log.loggingConfig import logger

# 重试
def retry_on_failure(max_retries=3,delay=5):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt < max_retries-1:
                        logger.warning(f"{func.__name__} 执行失败，准备重试")
                        logger.info(f"等待 {delay} 秒后进行第 {attempt+2} 次尝试...")
                        time.sleep(delay)
                    else:
                        logger.info(f"{func.__name__} 最终执行失败，已重试 {max_retries} 次")
                        raise
            return None
        return wrapper
    return decorator


# 登录JDE
@retry_on_failure(max_retries=3,delay=5)
def jde_login(in_page ,jde_url ,jde_environment ,jde_role ,jde_user ,jde_pass):

    in_page.goto(jde_url)

    element_exists = in_page.locator("#drop_mainmenu")
    if element_exists.is_visible(timeout=5000):
        logger.info("登录成功")
    else:
        # Details
        element_exists = in_page.get_by_role("textbox", name="Environment")
        if not element_exists.is_visible(timeout=2000):

            in_page.get_by_role("link", name="Details").click()
            in_page.wait_for_timeout(500)

        if jde_environment:
            in_page.get_by_role("textbox", name="Environment").fill(jde_environment)
            in_page.wait_for_timeout(500)
        if jde_role:
            in_page.get_by_role("textbox", name="Role").fill(jde_role)
            in_page.wait_for_timeout(500)

        # User Password
        in_page.get_by_role("textbox", name="User ID").fill(jde_user)
        in_page.wait_for_timeout(500)

        in_page.get_by_role("textbox", name="Password").fill(jde_pass)
        in_page.wait_for_timeout(500)

        # Sign In
        in_page.get_by_role("button", name="Sign In").click()
        in_page.wait_for_timeout(3000)

        element_exists = in_page.locator("#drop_mainmenu")
        element_exists.wait_for(state="visible", timeout=20000)
        if element_exists.is_visible():
            logger.info("登录成功")
        else:
            raise Exception("登录失败，未进入JDE界面，请检查！")




# 下载 GL movement
def download_Gl_movement(in_page):

    # 按科目帐的账户查询
    in_page.locator("#drop_mainmenu").click()
    in_page.get_by_role("cell", name="SWIRE PROPERTIES LTD - (CN)").first.click()
    in_page.get_by_role("cell", name="总帐").nth(1).click()




# 下载 TB
def download_TB():
    logger.info("")
