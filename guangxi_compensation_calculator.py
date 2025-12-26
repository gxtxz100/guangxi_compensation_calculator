#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
广西人身损害赔偿项目自动计算程序
根据广西最新标准计算各项赔偿项目并生成Word文档
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os


class GuangxiCompensationCalculator:
    """广西人身损害赔偿计算器"""
    
    # 2025年广西赔偿标准（示例数据，实际使用时需更新为最新标准）
    STANDARDS = {
        'urban_disposable_income': 45259,  # 城镇居民人均可支配收入（元/年）
        'rural_disposable_income': 19455,  # 农村居民人均纯收入（元/年）
        'urban_consumption': 29500,  # 城镇居民人均消费支出（元/年）
        'rural_consumption': 15400,  # 农村居民人均消费支出（元/年）
        'daily_meal_subsidy': 100,  # 住院伙食补助费（元/天）
        'daily_nursing_fee': 150,  # 护理费标准（元/天）
        'funeral_expense': 50000,  # 丧葬费（元）
    }
    
    # 伤残等级系数
    DISABILITY_COEFFICIENTS = {
        1: 1.0,
        2: 0.9,
        3: 0.8,
        4: 0.7,
        5: 0.6,
        6: 0.5,
        7: 0.4,
        8: 0.3,
        9: 0.2,
        10: 0.1
    }
    
    def __init__(self, root):
        self.root = root
        self.root.title("广西人身损害赔偿计算器")
        self.root.geometry("800x900")
        self.root.resizable(True, True)
        
        # 创建主框架
        self.create_widgets()
        
    def create_widgets(self):
        """创建GUI组件"""
        # 创建滚动框架
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 标题
        title_label = tk.Label(scrollable_frame, text="广西人身损害赔偿计算器", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 基本信息框架
        basic_frame = ttk.LabelFrame(scrollable_frame, text="基本信息", padding=10)
        basic_frame.pack(fill="x", padx=10, pady=5)
        
        self.victim_name = self.create_entry(basic_frame, "受害人姓名：", 0)
        self.victim_age = self.create_entry(basic_frame, "受害人年龄：", 1)
        self.victim_type = self.create_combobox(basic_frame, "户籍类型：", 
                                                 ["城镇", "农村"], 2)
        self.accident_date = self.create_entry(basic_frame, "事故发生日期（YYYY-MM-DD）：", 3)
        
        # 医疗相关费用框架
        medical_frame = ttk.LabelFrame(scrollable_frame, text="医疗相关费用", padding=10)
        medical_frame.pack(fill="x", padx=10, pady=5)
        
        self.medical_expense = self.create_entry(medical_frame, "医疗费（元）：", 0)
        self.hospital_days = self.create_entry(medical_frame, "住院天数：", 1)
        self.meal_subsidy = self.create_entry(medical_frame, "住院伙食补助费（元/天，默认100）：", 2)
        self.nutrition_fee = self.create_entry(medical_frame, "营养费（元）：", 3)
        self.traffic_fee = self.create_entry(medical_frame, "交通费（元）：", 4)
        self.accommodation_fee = self.create_entry(medical_frame, "住宿费（元）：", 5)
        
        # 误工费框架
        work_frame = ttk.LabelFrame(scrollable_frame, text="误工费", padding=10)
        work_frame.pack(fill="x", padx=10, pady=5)
        
        self.monthly_income = self.create_entry(work_frame, "月收入（元）：", 0)
        self.work_loss_days = self.create_entry(work_frame, "误工天数：", 1)
        
        # 护理费框架
        nursing_frame = ttk.LabelFrame(scrollable_frame, text="护理费", padding=10)
        nursing_frame.pack(fill="x", padx=10, pady=5)
        
        self.nursing_days = self.create_entry(nursing_frame, "护理天数：", 0)
        self.nursing_fee_per_day = self.create_entry(nursing_frame, "护理费标准（元/天，默认150）：", 1)
        self.nursing_count = self.create_entry(nursing_frame, "护理人数：", 2)
        
        # 残疾相关框架
        disability_frame = ttk.LabelFrame(scrollable_frame, text="残疾赔偿", padding=10)
        disability_frame.pack(fill="x", padx=10, pady=5)
        
        self.disability_level = self.create_combobox(disability_frame, "伤残等级：", 
                                                     [f"{i}级" for i in range(1, 11)], 0)
        self.disability_appliance_fee = self.create_entry(disability_frame, "残疾辅助器具费（元）：", 1)
        
        # 被扶养人生活费框架
        dependent_frame = ttk.LabelFrame(scrollable_frame, text="被扶养人生活费", padding=10)
        dependent_frame.pack(fill="x", padx=10, pady=5)
        
        self.dependent_count = self.create_entry(dependent_frame, "被扶养人数量：", 0)
        self.dependent_ages = self.create_entry(dependent_frame, "被扶养人年龄（用逗号分隔，如：5,10,15）：", 1)
        
        # 死亡相关框架
        death_frame = ttk.LabelFrame(scrollable_frame, text="死亡赔偿（如适用）", padding=10)
        death_frame.pack(fill="x", padx=10, pady=5)
        
        self.is_death = tk.BooleanVar()
        tk.Checkbutton(death_frame, text="是否死亡", variable=self.is_death).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        # 精神损害抚慰金框架
        mental_frame = ttk.LabelFrame(scrollable_frame, text="精神损害抚慰金", padding=10)
        mental_frame.pack(fill="x", padx=10, pady=5)
        
        self.mental_damage = self.create_entry(mental_frame, "精神损害抚慰金（元）：", 0)
        
        # 按钮框架
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill="x", padx=10, pady=20)
        
        calculate_btn = tk.Button(button_frame, text="计算赔偿", 
                                 command=self.calculate, bg="#4CAF50", 
                                 fg="white", font=("Arial", 12, "bold"),
                                 padx=20, pady=10)
        calculate_btn.pack(side="left", padx=5)
        
        export_btn = tk.Button(button_frame, text="导出Word文档", 
                               command=self.export_to_word, bg="#2196F3", 
                               fg="white", font=("Arial", 12, "bold"),
                               padx=20, pady=10)
        export_btn.pack(side="left", padx=5)
        
        clear_btn = tk.Button(button_frame, text="清空数据", 
                             command=self.clear_all, bg="#f44336", 
                             fg="white", font=("Arial", 12, "bold"),
                             padx=20, pady=10)
        clear_btn.pack(side="left", padx=5)
        
        # 结果显示框架
        result_frame = ttk.LabelFrame(scrollable_frame, text="计算结果", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.result_text = tk.Text(result_frame, height=15, wrap=tk.WORD, 
                                   font=("Arial", 10))
        self.result_text.pack(fill="both", expand=True)
        
        # 存储计算结果
        self.calculation_results = {}
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_entry(self, parent, label_text, row):
        """创建输入框"""
        tk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", padx=5, pady=5)
        entry = tk.Entry(parent, width=30)
        entry.grid(row=row, column=1, padx=5, pady=5)
        return entry
    
    def create_combobox(self, parent, label_text, values, row):
        """创建下拉框"""
        tk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", padx=5, pady=5)
        combobox = ttk.Combobox(parent, values=values, width=27, state="readonly")
        combobox.grid(row=row, column=1, padx=5, pady=5)
        if values:
            combobox.set(values[0])
        return combobox
    
    def get_float_value(self, entry, default=0.0):
        """获取浮点数值"""
        try:
            value = entry.get().strip()
            return float(value) if value else default
        except ValueError:
            return default
    
    def get_int_value(self, entry, default=0):
        """获取整数值"""
        try:
            value = entry.get().strip()
            return int(value) if value else default
        except ValueError:
            return default
    
    def calculate(self):
        """计算各项赔偿"""
        try:
            results = {}
            
            # 基本信息
            victim_name = self.victim_name.get().strip() or "未填写"
            victim_age = self.get_int_value(self.victim_age, 0)
            victim_type = self.victim_type.get()
            is_urban = (victim_type == "城镇")
            
            # 医疗费
            medical_expense = self.get_float_value(self.medical_expense)
            results['医疗费'] = medical_expense
            
            # 住院伙食补助费
            hospital_days = self.get_int_value(self.hospital_days)
            meal_subsidy_per_day = self.get_float_value(self.meal_subsidy, 
                                                       self.STANDARDS['daily_meal_subsidy'])
            meal_subsidy_total = hospital_days * meal_subsidy_per_day
            results['住院伙食补助费'] = meal_subsidy_total
            
            # 营养费
            nutrition_fee = self.get_float_value(self.nutrition_fee)
            results['营养费'] = nutrition_fee
            
            # 交通费
            traffic_fee = self.get_float_value(self.traffic_fee)
            results['交通费'] = traffic_fee
            
            # 住宿费
            accommodation_fee = self.get_float_value(self.accommodation_fee)
            results['住宿费'] = accommodation_fee
            
            # 误工费
            monthly_income = self.get_float_value(self.monthly_income)
            work_loss_days = self.get_int_value(self.work_loss_days)
            if monthly_income > 0 and work_loss_days > 0:
                daily_income = monthly_income / 30
                work_loss_fee = daily_income * work_loss_days
            else:
                work_loss_fee = 0
            results['误工费'] = work_loss_fee
            
            # 护理费
            nursing_days = self.get_int_value(self.nursing_days)
            nursing_count = self.get_int_value(self.nursing_count, 1)
            nursing_fee_per_day = self.get_float_value(self.nursing_fee_per_day, 
                                                       self.STANDARDS['daily_nursing_fee'])
            nursing_fee_total = nursing_days * nursing_fee_per_day * nursing_count
            results['护理费'] = nursing_fee_total
            
            # 残疾赔偿金
            disability_level_str = self.disability_level.get()
            if disability_level_str and disability_level_str != "无":
                disability_level = int(disability_level_str.replace("级", ""))
                coefficient = self.DISABILITY_COEFFICIENTS.get(disability_level, 0)
                base_income = (self.STANDARDS['urban_disposable_income'] if is_urban 
                              else self.STANDARDS['rural_disposable_income'])
                # 计算年限：75岁减去实际年龄，最低20年
                years = max(20, 75 - victim_age)
                disability_compensation = base_income * years * coefficient
                results['残疾赔偿金'] = disability_compensation
            else:
                results['残疾赔偿金'] = 0
            
            # 残疾辅助器具费
            disability_appliance_fee = self.get_float_value(self.disability_appliance_fee)
            results['残疾辅助器具费'] = disability_appliance_fee
            
            # 被扶养人生活费
            dependent_count = self.get_int_value(self.dependent_count)
            dependent_ages_str = self.dependent_ages.get().strip()
            dependent_living_expense = 0
            
            if dependent_count > 0 and dependent_ages_str:
                try:
                    ages = [int(age.strip()) for age in dependent_ages_str.split(',')]
                    base_consumption = (self.STANDARDS['urban_consumption'] if is_urban 
                                      else self.STANDARDS['rural_consumption'])
                    
                    for age in ages:
                        if age < 18:
                            years = 18 - age
                        elif age >= 60:
                            years = 20  # 通常按20年计算
                        else:
                            years = 0
                        
                        if years > 0:
                            # 按被扶养人数量分摊
                            expense = (base_consumption * years) / max(dependent_count, 1)
                            dependent_living_expense += expense
                except ValueError:
                    pass
            
            results['被扶养人生活费'] = dependent_living_expense
            
            # 死亡赔偿金
            if self.is_death.get():
                base_income = (self.STANDARDS['urban_disposable_income'] if is_urban 
                              else self.STANDARDS['rural_disposable_income'])
                years = max(20, 75 - victim_age)
                death_compensation = base_income * years
                results['死亡赔偿金'] = death_compensation
                results['丧葬费'] = self.STANDARDS['funeral_expense']
            else:
                results['死亡赔偿金'] = 0
                results['丧葬费'] = 0
            
            # 精神损害抚慰金
            mental_damage = self.get_float_value(self.mental_damage)
            results['精神损害抚慰金'] = mental_damage
            
            # 计算总计
            total = sum(results.values())
            results['总计'] = total
            
            # 保存结果
            self.calculation_results = results
            
            # 显示结果
            self.display_results(results, victim_name, victim_age, victim_type)
            
            messagebox.showinfo("成功", "计算完成！请查看计算结果。")
            
        except Exception as e:
            messagebox.showerror("错误", f"计算过程中出现错误：{str(e)}")
    
    def display_results(self, results, name, age, victim_type):
        """显示计算结果"""
        self.result_text.delete(1.0, tk.END)
        
        output = f"{'='*50}\n"
        output += f"广西人身损害赔偿计算结果\n"
        output += f"{'='*50}\n\n"
        output += f"受害人姓名：{name}\n"
        output += f"受害人年龄：{age}岁\n"
        output += f"户籍类型：{victim_type}\n"
        output += f"计算日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        output += f"\n{'-'*50}\n"
        output += f"各项赔偿明细：\n"
        output += f"{'-'*50}\n\n"
        
        for item, amount in results.items():
            if item != '总计':
                output += f"{item:20s}：{amount:>15,.2f} 元\n"
        
        output += f"\n{'-'*50}\n"
        output += f"{'总计':20s}：{results['总计']:>15,.2f} 元\n"
        output += f"{'='*50}\n"
        
        self.result_text.insert(1.0, output)
    
    def export_to_word(self):
        """导出到Word文档"""
        if not self.calculation_results:
            messagebox.showwarning("警告", "请先进行计算！")
            return
        
        try:
            # 选择保存位置
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
                initialfile=f"赔偿计算结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            )
            
            if not filename:
                return
            
            # 创建Word文档
            doc = Document()
            
            # 设置文档样式
            style = doc.styles['Normal']
            font = style.font
            font.name = '宋体'
            font.size = Pt(12)
            
            # 标题
            title = doc.add_heading('广西人身损害赔偿计算结果', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 基本信息
            doc.add_heading('一、基本信息', level=1)
            victim_name = self.victim_name.get().strip() or "未填写"
            victim_age = self.get_int_value(self.victim_age, 0)
            victim_type = self.victim_type.get()
            
            p = doc.add_paragraph()
            p.add_run('受害人姓名：').bold = True
            p.add_run(victim_name)
            
            p = doc.add_paragraph()
            p.add_run('受害人年龄：').bold = True
            p.add_run(f"{victim_age}岁")
            
            p = doc.add_paragraph()
            p.add_run('户籍类型：').bold = True
            p.add_run(victim_type)
            
            p = doc.add_paragraph()
            p.add_run('计算日期：').bold = True
            p.add_run(datetime.now().strftime('%Y年%m月%d日 %H:%M:%S'))
            
            # 赔偿明细
            doc.add_heading('二、赔偿明细', level=1)
            
            table = doc.add_table(rows=len(self.calculation_results), cols=2)
            table.style = 'Light Grid Accent 1'
            
            row_idx = 0
            for item, amount in self.calculation_results.items():
                if item == '总计':
                    continue
                table.rows[row_idx].cells[0].text = item
                table.rows[row_idx].cells[1].text = f"{amount:,.2f} 元"
                row_idx += 1
            
            # 总计
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.add_run('总计：').bold = True
            p.add_run(f"{self.calculation_results['总计']:,.2f} 元").bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # 计算依据
            doc.add_heading('三、计算依据', level=1)
            doc.add_paragraph('本计算依据《中华人民共和国民法典》及相关司法解释，')
            doc.add_paragraph('参考《2025年广西壮族自治区道路交通事故人身损害赔偿项目计算标准》')
            doc.add_paragraph('进行计算。')
            
            # 备注
            doc.add_heading('四、备注', level=1)
            doc.add_paragraph('1. 本计算结果仅供参考，实际赔偿金额以法院判决为准。')
            doc.add_paragraph('2. 各项费用需提供相应的票据和证明材料。')
            doc.add_paragraph('3. 如对计算结果有疑问，请咨询专业律师。')
            
            # 保存文档
            doc.save(filename)
            messagebox.showinfo("成功", f"Word文档已保存至：\n{filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出Word文档时出现错误：{str(e)}")
    
    def clear_all(self):
        """清空所有数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            # 清空所有输入框
            for widget in self.root.winfo_children():
                self._clear_widget(widget)
            
            self.result_text.delete(1.0, tk.END)
            self.calculation_results = {}
            messagebox.showinfo("提示", "数据已清空！")
    
    def _clear_widget(self, widget):
        """递归清空组件"""
        if isinstance(widget, tk.Entry):
            widget.delete(0, tk.END)
        elif isinstance(widget, ttk.Combobox):
            widget.set('')
        elif isinstance(widget, tk.Checkbutton):
            widget.deselect()
        elif hasattr(widget, 'winfo_children'):
            for child in widget.winfo_children():
                self._clear_widget(child)


def main():
    """主函数"""
    root = tk.Tk()
    app = GuangxiCompensationCalculator(root)
    root.mainloop()


if __name__ == "__main__":
    main()

