from logging import exception
from os.path import split
from os import remove
import pdfplumber
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog,messagebox
from PIL import Image,ImageTk

"""本文件为PDF_editor的主文件"""
class PDF_Page_Editor(Tk):
    def __init__(self,*args,**kwargs):
        super().__init__()
        self.title('PDF页面校正小工具@GAC ChengPu')
        #获取屏幕长宽，单位：像素
        self.screen_w = self.winfo_screenwidth()
        self.screen_h = self.winfo_screenheight()
        #计算窗口大小，宽度为屏幕宽度，高度需减去100
        self.window_w = self.screen_w
        self.window_h = self.screen_h-100
        #设定窗口大小，偏移0,0
        self.geometry(f"{self.window_w}x{self.window_h}+0+0")
        #初始化图片列表、文件路径
        self.images = []
        self.images_for_tk = []
        self.folder_path_str = ''
        self.file_name_str = ''
        self.current_image = None
        self.current_image_for_tk = None
        self.current_image_num = 0
        #初始化菜单栏，并显示窗口
        self.create_menu()

    def _quit(self):
        '''窗口关闭时，摧毁所有创建的进程'''
        self.destroy()

    def create_menu(self):
        '''创建页面的顶部菜单栏：文件操作，页面操作等'''
        #创建Menu框架，文件管理框架及工具框架
        menu_frame = Frame(self)
        menu_frame.grid(row=0,column=0,sticky=NW)
        menubar = Menu(menu_frame)
        # #canvas框架，画布框架
        # canvas_frame = Frame(self)
        # canvas_frame.grid(row=1,column=0)
        #创建文件操作菜单，注意command调用的函数不能加（）
        menubar.add_command(label='打开文件',command=self.open_pdf_file)
        menubar.add_command(label='添加文件',command=self.add_pdf_file)
        menubar.add_command(label='保存',command=self.save_pdf_file,accelerator='Ctrl+S')
        menubar.add_command(label='退出',command=self._quit)
        menubar.add_separator()
        menubar.add_separator()
        menubar.add_separator()
        menubar.add_command(label='删除本页',command=self.delete_page)
        menubar.add_command(label='旋转本页',command=self.rotate_page)
        menubar.add_command(label='本页前移',command=self.move_page_forward)
        menubar.add_command(label='本页后移',command=self.move_page_rear)
        menubar.add_separator()
        menubar.add_command(label='截取保留区域',command=self.cut_page)
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
        note_text = last_page_area.create_text(100,20,text='图片不清晰为显示问题',fill='white')
        last_page_area.create_text(100,40,text='导出PDF文件清晰度无影响',fill='white')
        last_page_area.create_text(100,60,text='显示清晰度与屏幕分辨率相关',fill='white')
        next_page_click_area = next_page_area.create_rectangle(0,0,int(self.window_w*0.1),int(self.window_h),fill="DeepSkyBlue",outline="DeepSkyBlue")
        next_page_text=next_page_area.create_text(int(self.window_w*0.05),int(self.window_h/2),text='下一页',fill='white',font=('Arial',20))
        #为上一页，下一页绑定鼠标键盘触发事件
        last_page_area.tag_bind(last_page_click_area,"<Button-1>",self.turn_to_last_page)
        next_page_area.tag_bind(next_page_click_area,"<Button-1>",self.turn_to_next_page)
        last_page_area.bind_all('<KeyPress-Left>',self.turn_to_last_page)
        next_page_area.bind_all('<KeyPress-Right>',self.turn_to_next_page)
        # 将图片显示到image_area区域
        if image :
            image_area.create_image(int(self.window_w/2),int(self.window_h),anchor='center',image=image)
            image_area.create_text(50,10,text=(f'第{self.current_image_num+1}页，共{len(self.images_for_tk)}页'),fill='black')
        else :
            print ('无图片输入')

    def add_pdf_file(self):
        """弹窗选取PDF文件,更新到class类数据中"""
        #弹窗选取文件，获取文件绝对路径
        self.file_path_str = filedialog.askopenfilename(initialdir=r'C:/',filetypes=[('PDF','pdf')])
        if self.file_path_str == '':
            print('文件未被打开')
        else:
            #拆分文件所在的路径与文件名称
            self.folder_path_str,self.file_name_str = split(self.file_path_str)
            #适用pdfplumber转换PDF文件为image对象
            try:
                with pdfplumber.open(self.file_path_str) as pdf:
                    # print (pdf)
                    images = []
                    for index,page in enumerate(pdf.pages):
                        #转换为图片，dpi设置为200
                        image = page.to_image(resolution=200)
                        #由于需要将pdfplumber格式图片转换为通用格式，这里采用笨办法，保存后重新打开
                        image_full_name = str(self.folder_path_str+'/pictures'+'_'+ str(index)+'.png')
                        image.save(image_full_name)
                        image = Image.open(image_full_name)
                        #重新加载完成后，调用os.remove删除本地临时保存的文件
                        # remove(image_full_name)
                        images.append(image)
            except Exception:
                messagebox.showerror(title='文件打开错误',message='文件可能被加密，无法打开...',)
                pass

        #打开文件后将图片转换为canvas格式，并调用load_image显示当前图片
        last_num = 0
        last_num = len (self.images_for_tk)
        for image in images:
            self.images.append(image)
            image_for_tk = ImageTk.PhotoImage(image)
            self.images_for_tk.append(image_for_tk)
        self.current_image_num = last_num
        self.current_image = self.images[last_num]
        self.current_image_for_tk = self.images_for_tk[last_num]
        self.load_image(self.images_for_tk[last_num])       

    def open_pdf_file(self):
        if len(self.images)>0:
            self.images = []
            self.images_for_tk = []
            self.current_image = None
            self.current_image_num = int()
        self.add_pdf_file()
        
    def save_pdf_file(self):
        pass

    def delete_page(self):
        pass

    def rotate_page(self):
        pass

    def move_page_forward(self):
        pass

    def move_page_rear(self):
        pass

    def cut_page(self):
        pass

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
        image_w,image_h = image.size
        w_scale = image_w/self.window_w
        h_scale = image_h/self.window_h
        scale = max(w_scale,h_scale)
        image_resized = image.resize((int(image_w/scale),int(image_h/scale)))
        return image_resized
    #

if __name__ == '__main__':
    App = PDF_Page_Editor()
    App.mainloop()
    #检测到窗口关闭，强制结束进程，必须在mainloop之后
    App.protocol("WM_DELETE_WINDOW",App._quit())