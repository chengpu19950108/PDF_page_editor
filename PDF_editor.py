"""
这是一个用于校正PDF文件页面顺序、方向，并进行简单合并、拆分、页面裁剪的小工具
需要安装第三方包：pillow/pdfplumber/img2pdf
使用中遇到的BUG及其他需求，欢迎联络：GAC R&D-程浦：chengpu@gacrnd.com
如果有非常复杂的需求，建议使用具有编辑功能的专业阅读器
"""
from os import makedirs,path
import pdfplumber
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog,messagebox
from PIL import Image,ImageTk
import img2pdf
import shutil

"""本文件为PDF_editor的主文件"""
class PDF_Page_Editor(Tk):
    def __init__(self,*args,**kwargs):
        super().__init__()
        self.title('PDF页面校正小工具@GAC_ChengPu_05921')
        #获取屏幕长宽，单位：像素
        self.screen_w = self.winfo_screenwidth()
        self.screen_h = self.winfo_screenheight()
        #计算窗口大小，宽度为屏幕宽度，高度需减去100
        self.window_w = self.screen_w
        self.window_h = self.screen_h-100
        #设定窗口大小，偏移0,0
        self.geometry(f"{self.window_w}x{self.window_h}+0+0")
        #初始化图片列表、文件路径
        self.image_num = 0                  #已打开的总的页面数量
        self.images = []                    #pillow图片列表
        self.images_for_tk = []             #tk格式图片列表
        self.original_image_folder = ''     #转换前的临时图片存储路径，文件保存后将被自动删除
        self.converted_image_folder = ''    #转换后的临时图片存储路径，文件保存后将被自动删除
        self.folder_path_str = ''           #当前打开的PDF路径
        self.file_name_str = ''             #当前打开的PDF文件名称
        self.current_image = None           #当前显示的图片原格式：pillow
        self.current_image_for_tk = None    #当前显示的图片处理为tk格式
        self.current_image_num = 0          #当前显示的图片索引号
        self.scale = []                     #图片的缩放比例列表
        self.rotation_angle = []            #图片的旋转角度列表
        self.cut_y = []                     #图片的裁剪参数列表
        #初始化菜单栏，并显示窗口
        self.create_menu()
        self.image_area = None
        self.cover_page_act = False         #初始化当前插入色块功能状态，初始不激活
        

    def _quit(self):
        '''窗口关闭时，清除变量、删除临时文件、摧毁所有创建的进程'''
        self.clear_temp()
        try:
            self.destroy()
        except Exception:
            pass

    def create_menu(self):
        '''创建页面的顶部菜单栏：文件操作，页面操作等'''
        #创建Menu框架，文件管理框架及工具框架
        menu_frame = Frame(self)
        menu_frame.grid(row=0,column=0,sticky=NW)
        menubar = Menu(menu_frame)
        #创建文件操作菜单，注意command调用的函数不能加（）
        menubar.add_command(label='打开文件',command=self.open_pdf_file)
        menubar.add_command(label='添加文件',command=self.add_pdf_file)
        menubar.add_command(label='保存并关闭',command=self.save_pdf_file,accelerator='Ctrl+S')
        menubar.add_command(label='退出',command=self._quit)
        menubar.add_separator()
        menubar.add_separator()
        menubar.add_separator()
        menubar.add_command(label='删除本页',command=self.delete_page)
        menubar.add_command(label='旋转本页',command=self.rotate_page)
        menubar.add_command(label='本页前移',command=self.move_page_forward)
        menubar.add_command(label='本页后移',command=self.move_page_rear)
        menubar.add_separator()
        cover_button = menubar.add_command(label='插入色块（遮盖不相关区域）',command=self.cover_page)
        self.config(menu=menubar)
          
    def load_image(self,image):
        '''创建画布区域，并显示传入的单张图片'''
        #canvas框架，画布框架
        canvas_frame = Frame(self)
        canvas_frame.grid(row=1,column=0)
        #创建分框架，左侧为上一页点击区域，中间为图片显示区域，右侧为下一页点击区域
        image_area = Canvas(canvas_frame,width=int(self.window_w*0.8),height=self.window_h,bg='white')
        last_page_area = Canvas(canvas_frame,width=(int(self.window_w*0.1)),height=self.window_h,bg='DeepSkyBlue')
        next_page_area = Canvas(canvas_frame,width=(int(self.window_w*0.1)),height=self.window_h,bg='DeepSkyBlue')
        last_page_area.grid(row=0,column=0)
        image_area.grid(row=0,column=1)
        next_page_area.grid(row=0,column=2)
        #填充左右侧画布区域，显示文字与颜色块
        last_page_click_area = last_page_area.create_rectangle(0,0,int(self.window_w*0.1),int(self.window_h),fill="DeepSkyBlue",outline="DeepSkyBlue")
        last_page_text=last_page_area.create_text(int(self.window_w*0.05),int(self.window_h/2),text='上一页',fill='white',font=('Arial',20))
        next_page_click_area = next_page_area.create_rectangle(0,0,int(self.window_w*0.1),int(self.window_h),fill="DeepSkyBlue",outline="DeepSkyBlue")
        next_page_text=next_page_area.create_text(int(self.window_w*0.05),int(self.window_h/2),text='下一页',fill='white',font=('Arial',20))
        #为上一页，下一页绑定鼠标键盘触发事件
        last_page_area.tag_bind(last_page_click_area,"<Button-1>",self.turn_to_last_page)
        next_page_area.tag_bind(next_page_click_area,"<Button-1>",self.turn_to_next_page)
        last_page_area.tag_bind(last_page_text,"<Button-1>",self.turn_to_last_page)
        next_page_area.tag_bind(next_page_text,"<Button-1>",self.turn_to_next_page)
        last_page_area.bind_all('<KeyPress-Left>',self.turn_to_last_page)
        next_page_area.bind_all('<KeyPress-Right>',self.turn_to_next_page)
        # 将图片显示到image_area区域
        if image :
            image_area.create_image(int(self.window_w*0.8/2),int(self.window_h/2),anchor='center',image=image)
            image_area.create_text(50,10,text=(f'第{self.current_image_num+1}页，共{len(self.images_for_tk)}页'),fill='black')
            # image_area.create_rectangle(100,0,1000,800,outline='black')
        else :
            print ('当前无输入')
        #返回一个canvas对象，方便后续在本区域绘制交互图形
        return (image_area)

    def add_pdf_file(self):
        """弹窗选取PDF文件,更新到类变量中。不清除已打开的文件，将新文件自动添加到队列后，可实现合并pdf文件的功能"""
        #弹窗选取文件，获取文件绝对路径
        self.file_path_str = filedialog.askopenfilename(initialdir='Desktop',filetypes=[('PDF','pdf')])
        #文件打开后，校验文件是否正确、临时文件夹是否正确、队列中是否已有加载的文件，
        if self.file_path_str == '':
            print('文件未被打开')
        else:
            #拆分文件所在的路径与文件名称
            folder_path,file_name = path.split(self.file_path_str)
            if self.folder_path_str == '':
                self.folder_path_str = folder_path
                if self.file_name_str == '':
                    self.file_name_str = file_name
                else:
                    pass
            else:
                pass
            #校验临时文件夹是否正确
            if self.original_image_folder == '' :
                temp_path_original = path.join(self.folder_path_str,'original_images')
                temp_path_converted = path.join(self.folder_path_str,'converted_images')
                #校验是否已存在原始图片的临时文件夹
                if path.exists(temp_path_original):
                    shutil.rmtree(temp_path_original,ignore_errors=True)
                makedirs(temp_path_original)
                self.original_image_folder =temp_path_original
                #校验是否已存在转换图片的临时文件夹
                if path.exists(temp_path_converted):
                    shutil.rmtree(temp_path_converted,ignore_errors=True)
                makedirs(temp_path_converted)
                self.converted_image_folder =temp_path_converted
            else:
                pass
            #转换PDF文件为适用pdfplumberimage对象
            try:
                with pdfplumber.open(self.file_path_str) as pdf:
                    images = []
                    for index,page in enumerate(pdf.pages):
                        #转换为图片，dpi设置为200
                        image = page.to_image(resolution=200)
                        #由于需要将pdfplumber格式图片转换为通用格式，这里采用笨办法，保存后重新打开
                        image_full_name = path.join(self.original_image_folder,'image_'+ str(self.image_num+index)+'.png')
                        image.save(image_full_name)
                        image = Image.open(image_full_name)
                        images.append(image)
            except Exception:
                messagebox.showerror(title='文件打开错误',message='文件可能被加密，无法打开...',)
                return
        #打开文件完成后，将原图转换为tk格式，并调用load_image显示当前图片
        for image in images:
            #原始图片添加到队列
            self.images.append(image)
            #尺寸缩放后，缩放比例添加到队列，缩放比例并转换TK格式的图片添加到队列
            scale,image_scaled = self.resize_image(image)
            image_for_tk = ImageTk.PhotoImage(image_scaled)
            self.scale.append(scale)
            self.rotation_angle.append(0)
            # self.cut_y.append(0)
            self.images_for_tk.append(image_for_tk)
        self.current_image_num = self.image_num
        self.current_image = self.images[self.image_num]
        self.current_image_for_tk = self.images_for_tk[self.image_num]
        self.image_num = len (self.images)
        #加载完文件后，初始化旋转角度列表、裁剪清单列表
        self.image_area = self.load_image(self.current_image_for_tk)
        # image_area.create_rectangle(100,0,1000,800,outline='black')

    def open_pdf_file(self):
        """清除已有变量值后打开新的工作区"""
        if len(self.images)>0:
            self.clear_vars()
            self.clear_temp()
        #参数全部初始化后，执行添加文件操作
        self.add_pdf_file()
        
    def save_pdf_file(self):
        """将当前进程中校正后的所有页面保存为一个PDF文件，存放在第一个打开的文件所在的路径"""
        if self.images:
            image_path_list = []
            for index,image in enumerate(self.images):
                image_path = path.join(self.converted_image_folder,str('image_'+str(index)+'.png'))
                #对image进行旋转、裁剪等操作后
                image.save(image_path)
                image_path_list.append(image_path)
                # self.images = image.cut(self.cut_y[index])
            output_file_name = path.join(self.folder_path_str,str('converted_'+self.file_name_str))
            with open (output_file_name,'wb') as f:
                write_content = img2pdf.convert(image_path_list)
                f.write(write_content)
            messagebox.showinfo(title='成功',message=f'yeah~文件转换成功。已保存至原路径，详细路径：{output_file_name}')
            #完成保存任务后，删除所有文件夹
            if path.exists(self.original_image_folder):
                shutil.rmtree(self.original_image_folder,ignore_errors=True)
            if path.exists(self.converted_image_folder):
                shutil.rmtree(self.converted_image_folder,ignore_errors=True)
            self.clear_vars()
            self.load_image(image=None)


        else:
            pass

    def delete_page(self):
        n = self.current_image_num
        m = self.image_num
        #只有一页或无页面的情况下，不能删除
        if m == 1:
            messagebox.showinfo(title='无法删除',message='这个文档只有一页了，无法继续删除...')
            return
        if m == 0:
            return
        #页面数大于1时，删除各队列中的本页相关内容
        self.image_num -= 1
        del self.images[n]
        del self.images_for_tk[n]
        del self.scale[n]
        del self.rotation_angle[n]
        # del self.cut_y[n]
        #重新定位当前页面：删除了最后一页时，定位到首页，否则定位+1
        if n+1 == m:
            new_n = 0
            self.current_image_num = new_n
        else:
            new_n = n
            self.current_image_num = new_n
        #更新当前页面的信息
        self.current_image = self.images [new_n]
        self.current_image_for_tk = self.images_for_tk[new_n]
        #序列更新完成后加载页面
        self.image_area = self.load_image(self.current_image_for_tk)

    def rotate_page(self):
        """对当前pillow格式图片逆时针旋转90度，并将更新到角度队列"""
        n = self.current_image_num
        self.images[n] = self.current_image.rotate(90,expand=True)
        self.current_image = self.images[n]
        self.rotation_angle[n] += 90
        #旋转完成后重新计算长宽比例，load图像
        self.scale[n],image_scaled = self.resize_image(self.current_image)
        self.images_for_tk[n] = ImageTk.PhotoImage(image_scaled)
        self.current_image_for_tk = self.images_for_tk[n]
        self.load_image(self.current_image_for_tk)
    
    def move_page_forward(self):
        n = self.current_image_num
        m = self.image_num
        #只有一页或无页面的情况下，不能移动
        if m == 1 or m == 0:
            return
        if n == 0:
            messagebox.showinfo(title='无法移动',message='已经移动到首页了，不能继续前移了...')
            return
        else:
            new_n = n-1
            self.current_image_num = new_n
            image = self.images.pop(n)
            self.images.insert(new_n,image)
            image_for_tk = self.images_for_tk.pop(n)
            self.images_for_tk.insert(new_n,image_for_tk)
            scale = self.scale.pop(n)
            self.scale.insert(new_n,scale)
            rotation = self.rotation_angle.pop(n)
            self.rotation_angle.insert(new_n,rotation)
            # cut_y = self.cut_y.pop(n)
            # self.cut_y.insert(new_n,cut_y)
        #队列更新完成后，重新load图像
        self.image_area = self.load_image(self.current_image_for_tk)

    def move_page_rear(self):
        n = self.current_image_num
        m = self.image_num
        #只有一页或无页面的情况下，不能移动
        if m == 1 or m == 0:
            return
        if n+1 == m:
            messagebox.showinfo(title='无法移动',message='已经移动到尾页了，不能继续后移了...')
            return
        else:
            new_n = n+1
            self.current_image_num = new_n
            image = self.images.pop(n)
            self.images.insert(new_n,image)
            image_for_tk = self.images_for_tk.pop(n)
            self.images_for_tk.insert(new_n,image_for_tk)
            scale = self.scale.pop(n)
            self.scale.insert(new_n,scale)
            rotation = self.rotation_angle.pop(n)
            self.rotation_angle.insert(new_n,rotation)
            # cut_y = self.cut_y.pop(n)
            # self.cut_y.insert(new_n,cut_y)
        #队列更新完成后，重新load图像
        self.image_area = self.load_image(self.current_image_for_tk)

    def cover_page(self):
        """插入一个色块，用于遮盖不相关区域达到保密的目的"""
        
        if self.cover_page_act == True:
            self.cover_page_act = False
            self.image_area = self.load_image(self.current_image_for_tk)
            return
        else:
            self.cover_page_act = True
            n = self.current_image_num
            image = self.current_image
            scale = self.scale[n]
            self.image_area.bind ("<Button-1>",self.on_left_button_click)
            self.image_area.bind ("<B1-Motion>",self.on_mouse_drag)
            self.image_area.configure(cursor = "cross")
            self.start_x = None
            self.start_y = None
        
    def on_left_button_click(self,event):
        self.start_x = self.image_area.canvasx(event.x)
        self.start_y = self.image_area.canvasy(event.y)

    def on_mouse_drag(self,event):
        x1, y1 = self.image_area.canvasx(event.x), self.image_area.canvasy(event.y)
        x2, y2 = self.start_x, self.start_y
        self.image_area.delete("rect")
        self.image_area.create_rectangle(x1, y1, x2, y2, tags="rect", outline="red")
        self.start_x, self.start_y = x2, y2

    # def change_cursor(self,event):
    #     event.widget.cursor("cross")

    def turn_to_last_page(self,event):
        '''将当前页面数-1，如果已经为0，那么设置为最大值，并更新当前image 和 imgae_for_tk'''
        if self.current_image_num == 0:
            self.current_image_num = len(self.images)-1
        else:
            self.current_image_num -= 1
        self.current_image = self.images[self.current_image_num]
        self.current_image_for_tk = self.images_for_tk[self.current_image_num]
        #重设图片大小并显示
        #self.resize()
        self.load_image(self.current_image_for_tk)
        
    def turn_to_next_page(self,event):
        '''将当前页面数+1，如果已经为最大值，那么设置为0，并更新当前image 和 imgae_for_tk'''
        if self.current_image_num == len(self.images)-1:
            self.current_image_num = 0
        else:
            self.current_image_num += 1
        self.current_image = self.images[self.current_image_num]
        self.current_image_for_tk = self.images_for_tk[self.current_image_num]
        #重设图片大小并显示
        #self.resize()
        self.load_image(self.current_image_for_tk)

    def resize_image(self,image):
        """根据初始窗口尺寸和图片原始尺寸，重新调整图片以适应窗口大小，传入的图片为pillow格式"""
        image_w,image_h = image.size
        w_scale = image_w/self.window_w
        h_scale = image_h/self.window_h
        scale = max(w_scale,h_scale)
        #先转换为黑白图片显示，以保证后面图片处理后的清晰度
        image = image.convert('L')
        image_resized = image.resize((int(image_w/scale),int(image_h/scale)),resample=Image.BICUBIC)
        return (scale,image_resized)
    
    def clear_vars(self):
        """将进程中的所有变量初始化"""
        self.image_num = 0
        self.scale = []
        self.images = []
        self.images_for_tk = []
        self.current_image = None
        self.current_image_for_tk = None
        self.current_image_num = 0
        self.original_image_folder = ''
        self.converted_image_folder = ''
        self.folder_path_str = ''
        self.file_name_str = ''
        self.rotation_angle = []
        self.cut_y = []

    def clear_temp(self):
        """清除本地保存的临时文件，包含一个原图片文件夹和一个转换后的图片文件夹"""
        if path.exists(self.original_image_folder):
            shutil.rmtree(self.original_image_folder,ignore_errors=True)
        if path.exists(self.converted_image_folder):
            shutil.rmtree(self.converted_image_folder,ignore_errors=True)

if __name__ == '__main__':
    App = PDF_Page_Editor()
    App.mainloop()
    #检测到窗口关闭时强制结束进程，必须在mainloop之后
    try:
        App.protocol("WM_DELETE_WINDOW",App._quit())
    except Exception:
        pass