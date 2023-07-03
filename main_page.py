from tkinter import *
from tkinter import ttk
from tkinter import messagebox
# 导入图形库
from PIL import ImageTk, Image
# 导入全局格式数据
from quanju_data_set import *
# 导入生成公司数据的文件
from create_new_company_data import write_save_all_company
# 导入在线数据
from get_new_data import get_new_data
# 导入转换并合并成PDF的程序
from combine_docx_to_one_pdf import combine_process
# 导入修改证书的程序
from change_brand_broken import change_all_brand, change_all_problem_file

# 获取在线数据
product_json, factory_list = get_new_data()


# 创建首页
def create_main_page(root):
    # 创建首页
    main_frame = Frame(root, width=MAIN_WIDTH, height=MAIN_HEIGHT, bg=BG_COLOR)
    main_frame.pack_propagate(False)

    # 设置首页标题
    title_label = Label(main_frame, text="越鑫检测证书生成", bg=BG_COLOR, fg="white",
                        font=TITLE_FONT)
    title_label.pack(pady=30)

    # 设置证书生成的按钮
    photo = ImageTk.PhotoImage(Image.open("images/make.png").resize((20, 20)))
    make_button = Button(main_frame,
                         text="证书生成",
                         image=photo,
                         compound='left',
                         font=("Ubuntu", 16),
                         bg=BG_COLOR,
                         fg="white",
                         command=lambda: display_file_page(root, main_frame)
                         )
    make_button.image = photo
    make_button.configure(padx=10)
    make_button.pack(ipadx=25, ipady=10, pady=20)

    # 设置证书修改的按钮
    photo2 = ImageTk.PhotoImage(Image.open("images/change.png").resize((20, 20)))
    change_button = Button(main_frame,
                           text="证书修改",
                           image=photo2,
                           compound='left',
                           font=("Ubuntu", 16),
                           bg=BG_COLOR,
                           fg="white",
                           command=lambda: display_edit_file_page(root, main_frame)
                           )
    change_button.image = photo2
    change_button.configure(padx=10)
    change_button.pack(ipadx=25, ipady=10, pady=20)

    info = Label(main_frame,
                 text="---------使用流程--------\n\n（1）生成证书\n\n（2）修改品牌\n\n（3）修改故障\n\n（4）合并证书",
                 bg=BG_COLOR, fg="white",
                 font=("Ubuntu", 12))
    info.pack(pady=10)

    return main_frame


# 显示首页
def display_main_page(root, other_page=None):
    # 如果是第一次调用，就直接加载首页
    main_page = create_main_page(root)
    main_page.pack(fill=BOTH, expand=True)

    # 如果是从其他页面调用，则加载后需要将其他页面隐藏，并调整窗口
    if other_page is not None:
        other_page.pack_forget()
        root.geometry(f"{MAIN_WIDTH}x{MAIN_HEIGHT}")
        root.config(padx=0, pady=0)


# 创建证书生成页面
def create_file_page(root):
    def main_process():
        show_sections = sections_input.get().split(' ')
        show_sections_num_input = sections_num_input.get().split(' ')
        if messagebox.askyesno(title="输入信息确认",
                               message=f"公司名称：{company_input.get()}\n\n"
                                       f"探头数量：{alert_num_input.get()}\n\n"
                                       f"检测日期：{test_date_input.get()}\n\n"
                                       f"温度：{temperature_input.get()}\n\n"
                                       f"湿度：{humidity_input.get()}\n\n"
                                       f"区域分布：{', '.join(show_sections)}\n\n"
                                       f"各区域探头数量：{', '.join(show_sections_num_input)}\n\n"
                                       f"证书起始编号：{start_num_input.get()}\n\n"
                                       f"请核对以上输入的信息，点击'否'返回修改，点击'是'继续生成证书"):
            user_input = {
                "company_name": company_input.get(),
                "all_nums": int(alert_num_input.get()),
                "date": test_date_input.get(),
                "temperature": temperature_input.get(),
                "humidity": humidity_input.get(),
                "sections": sections_input.get().strip().split(' '),
                "sections_num": [int(num) for num in sections_num_input.get().strip().split(' ')],
                "start_num": int(start_num_input.get())
            }
            write_save_all_company(user_input)
            messagebox.showinfo(title="报告生成", message="全部的报告已经生成，请到指定的公司文件夹查看")

    root.geometry("650x820")
    root.config(padx=50, pady=20)
    make_new_files_frame = Frame(root)

    title_label = Label(make_new_files_frame, text="越鑫检测证书生成器", pady=15, font=TITLE_FONT)
    title_label.grid(row=0, column=0, columnspan=2)
    company_label = Label(make_new_files_frame, text="公司名称：", pady=18, font=CONTEXT_FONT, height=2)
    company_label.grid(row=1, column=0, sticky="E")
    company_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    company_input.grid(row=1, column=1, ipady=5)
    alert_num_label = Label(make_new_files_frame, text="探头总数量：", pady=18, font=CONTEXT_FONT, height=2)
    alert_num_label.grid(row=2, column=0, sticky="E")
    alert_num_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    alert_num_input.grid(row=2, column=1, ipady=5)
    test_date_label = Label(make_new_files_frame, text="检测日期（格式20220116）：", pady=18, font=CONTEXT_FONT,
                            height=2)
    test_date_label.grid(row=3, column=0, sticky="E")
    test_date_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    test_date_input.grid(row=3, column=1, ipady=5)
    temperature_label = Label(make_new_files_frame, text="温度（数值）：", pady=18, font=CONTEXT_FONT, height=2)
    temperature_label.grid(row=4, column=0, sticky="E")
    temperature_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    temperature_input.grid(row=4, column=1, ipady=5)
    humidity_label = Label(make_new_files_frame, text="湿度（数值）：", pady=18, font=CONTEXT_FONT, height=2)
    humidity_label.grid(row=5, column=0, sticky="E")
    humidity_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    humidity_input.grid(row=5, column=1, ipady=5)
    sections_label = Label(make_new_files_frame, text="探头分布区域，以空格分隔：", pady=18, font=CONTEXT_FONT)
    sections_label.grid(row=6, column=0, sticky="E")
    sections_input = Entry(make_new_files_frame, width=47, font=CONTEXT_FONT_2)
    sections_input.grid(row=6, column=1, ipady=5)
    sections_num_label = Label(make_new_files_frame, text="各区域的探头数量，以空格分隔：", pady=18, font=CONTEXT_FONT)
    sections_num_label.grid(row=7, column=0, sticky="E")
    sections_num_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    sections_num_input.grid(row=7, column=1, ipady=5)
    start_num = Label(make_new_files_frame, text="证书起始编号：", pady=18, font=CONTEXT_FONT)
    start_num.grid(row=8, column=0, sticky="E")
    start_num_input = Entry(make_new_files_frame, width=35, font=CONTEXT_FONT)
    start_num_input.insert(0, "1")
    start_num_input.grid(row=8, column=1, ipady=5)

    new_image = ImageTk.PhotoImage(Image.open("images/new.png").resize((20, 20)))
    submit = Button(make_new_files_frame, text="生成证书", image=new_image, compound='left',
                    font=("Microsoft YaHei", 12, "bold"), command=main_process)
    submit.image = new_image
    submit.configure(padx=10)
    submit.grid(row=9, column=0, columnspan=2, pady=25, ipadx=20, ipady=10)

    photo2 = ImageTk.PhotoImage(Image.open("images/home.png").resize((20, 20)))
    change_button = Button(make_new_files_frame,
                           text="返回首页",
                           image=photo2,
                           compound='left',
                           font=SUBMIT_FONT,
                           bg="#678983",
                           fg="white",
                           command=lambda: display_main_page(root, other_page=make_new_files_frame)
                           )
    change_button.image = photo2
    change_button.configure(padx=10)
    change_button.grid(row=10, column=0, columnspan=2, ipadx=20, ipady=10)

    return make_new_files_frame


# 显示证书生成页面
def display_file_page(root, main_frame):
    main_frame.pack_forget()
    create_file_page(root).pack(fill=BOTH, expand=True)


# 创建证书修改页面
def edit_file_page(root):
    root.geometry("550x520")
    root.config(padx=50, pady=20)
    change_files_frame = Frame(root)

    change_company_label = Label(change_files_frame, text="公司名称：", pady=18, font=CONTEXT_FONT)
    change_company_label.grid(row=1, column=0, sticky='E', ipadx=10)
    change_company_input = Entry(change_files_frame, width=30, font=CONTEXT_FONT)
    change_company_input.grid(row=1, column=1, ipady=4)
    change_sections_num_label = Label(change_files_frame, text="编号：", pady=28, font=CONTEXT_FONT)
    change_sections_num_label.grid(row=2, column=0, sticky='E', ipadx=10)
    change_sections_num_input = Text(change_files_frame, width=30, height=2, font=CONTEXT_FONT, spacing1=8)
    change_sections_num_input.grid(row=2, column=1)
    Label(change_files_frame, text='编号以空格分割，示例：1-5 6 8', fg='gray').grid(row=3, column=1, sticky='W')

    # 品牌选中时的处理函数
    def selected(self):
        type_choice.config(values=[item['list'] for item in product_json if item['name'] == factory_choice.get()][0])
        type_choice.current(0)

    factory_label = Label(change_files_frame, text="品牌:", pady=10, font=CONTEXT_FONT)
    factory_label.grid(row=6, column=0, sticky='E', ipadx=10)
    factory_choice = ttk.Combobox(change_files_frame, width=28, state="readonly", font=CONTEXT_FONT)
    factory_choice.bind("<<ComboboxSelected>>", selected)
    factory_choice['values'] = factory_list
    root.option_add('*TCombobox*Listbox.font', CONTEXT_FONT)
    factory_choice.grid(row=6, column=1)
    factory_choice.current(0)

    type_label = Label(change_files_frame, text="型号:", pady=20, font=CONTEXT_FONT)
    type_label.grid(row=7, column=0, sticky='E', ipadx=10)
    type_choice = ttk.Combobox(change_files_frame, width=28, height=22, state="readonly", font=CONTEXT_FONT)
    type_choice['values'] = [item['list'] for item in product_json if item['name'] == factory_choice.get()][0]
    root.option_add('*TCombobox*Listbox.font', CONTEXT_FONT)
    type_choice.grid(row=7, column=1)
    type_choice.current(0)

    ttk.Separator(change_files_frame, orient='horizontal').grid(row=9, columnspan=4, sticky='ew')

    break_image = ImageTk.PhotoImage(Image.open("images/break.png").resize((20, 20)))
    break_button = Button(change_files_frame, text="故障探头", image=break_image, compound='left', font=SUBMIT_FONT,
                          command=lambda: change_all_problem_file(change_company_input.get(),
                                                                  change_sections_num_input.get("1.0", END))
                          )
    break_button.image = break_image
    break_button.configure(padx=10)
    break_button.grid(row=10, column=0, pady=30, ipadx=29, ipady=13, sticky='E')

    brand_image = ImageTk.PhotoImage(Image.open("images/brand.png").resize((20, 20)))
    brand_button = Button(change_files_frame, image=brand_image, compound='left', text="修改品牌",
                          command=lambda: change_all_brand(change_company_input.get(), factory_choice.get(),
                                                           type_choice.get(),
                                                           change_sections_num_input.get("1.0", END)),
                          font=SUBMIT_FONT)
    brand_button.image = brand_image
    brand_button.configure(padx=10)
    brand_button.grid(row=10, column=1, ipadx=29, ipady=13, sticky='E')

    combine_image = ImageTk.PhotoImage(Image.open("images/combine.png").resize((20, 20)))
    combine_button = Button(change_files_frame, text="证书合并", image=combine_image,
                            compound='left',
                            font=SUBMIT_FONT,
                            command=lambda: combine_process(change_company_input.get()))
    combine_button.image = combine_image
    combine_button.configure(padx=10)
    combine_button.grid(row=13, column=0, ipadx=29, ipady=13, sticky='E')

    photo2 = ImageTk.PhotoImage(Image.open("images/home.png").resize((20, 20)))
    back_button = Button(change_files_frame,
                         text="返回首页",
                         image=photo2,
                         compound='left',
                         font=SUBMIT_FONT,
                         bg="#678983",
                         fg="white",
                         command=lambda: display_main_page(root, other_page=change_files_frame))
    back_button.image = photo2
    back_button.configure(padx=10)
    back_button.grid(row=13, column=1, ipadx=29, ipady=13, sticky='E')

    return change_files_frame


# 显示证书生成emm
def display_edit_file_page(root, main_frame):
    main_frame.pack_forget()
    edit_file_page(root).pack(fill=BOTH, expand=True)


def display_main():
    root = Tk()
    root.title("越鑫检测证书生成")
    root.resizable(False, False)  # 固定窗口大小
    root.config(padx=0, pady=0)

    display_main_page(root)

    root.mainloop()


if __name__ == "__main__":
    display_main()
