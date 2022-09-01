classdef MathCode < handle

    properties (Access = public)
        originData%原始数据
        data%读取预测(计算)数据 *
        windloss
        loadloss
        out

        loadrate%导入负荷功率比例    *
        windrate%导入风电功率比例    *

        loadmax=900;%设置负荷最大值MW
        windmax%设置风电装机容量(300/600/900)MW    *

        ld%负荷功率   *
        wd%风电功率   *

        unit1
        unit2
        unit3
        %机组(该值为1时为机组开启,为0时为关闭)  *

        carbon%碳捕集成本

        store%进行储能(1为储能0为不储能)    *
        store1%与store相反  *

        wdtotal%总弃风  *
        ldtotal%总失负荷 *

        machine1% 1:机组1功率
        machine2% 2:机组2功率
        machine3% 3:机组3功率
        totalcost% 4:总成本    *
        abwindcost% 5:弃风损失  *
        loadcost% 6:失负荷损失   *
        carboncost% 7:碳捕集成本 *
        storecost% 8:储能成本   *
        windcost% 9:风电成本    *
        firecost% 10:火电成本   *
        unitprive% 11:单价    *

        rowOfData2=1%数据起始行为1
        %机组最大技术出力
        mx1=600
        mx2=300
        mx3=150
        %机组最小技术出力
        mn1=180
        mn2=90
        mn3=45
        %碳排放量
        ce1=0.72
        ce2=0.75
        ce3=0.79
        %煤耗量系数c
        cc1=786.80
        cc2=451.32
        cc3=1049.50
        %煤耗量系数b
        cb1=30.42
        cb2=65.12
        cb3=139.6
        %煤耗量系数a
        ca1=0.226
        ca2=0.588
        ca3=0.785

    end

    methods
        function obj = MathCode(unit1,unit2,unit3,carboon,store,store1,windmax)
            obj.unit1 = unit1;
            obj.unit2 = unit2;
            obj.unit3 = unit3;
            obj.carbon = carboon;
            obj.store = store;
            obj.store1 = store1;
            obj.windmax = windmax;
        end
        function GetData(obj,data)
            obj.originData=xlsread(data);
            if data == "预测数据表2.xlsx"
                obj.data = obj.originData(obj.rowOfData2:obj.rowOfData2+95,1:3);
            else
                obj.data = obj.originData;
            end
        end

        function f1(obj)
            obj.loadrate=obj.data(:,2);
            obj.windrate=obj.data(:,3);
            obj.windloss=ones(96,1)*0;
            obj.loadloss=ones(96,1)*0;
            obj.ld=obj.loadrate*obj.loadmax;
            obj.wd=obj.windrate*obj.windmax;
            for i=1:1:96
                if obj.wd(i)+obj.mn1*obj.unit1+obj.mn2*obj.unit2+obj.mn3*obj.unit3-obj.ld(i)>=0
                    obj.windloss(i)=obj.wd(i)+obj.mn1*obj.unit1+obj.mn2*obj.unit2+obj.mn3*obj.unit3-obj.ld(i);
                else
                    obj.windloss(i)=0;
                end%弃风功率
                if obj.ld(i)-obj.wd(i)-obj.mx1*obj.unit1-obj.mx2*obj.unit2-obj.mx3*obj.unit3>=0
                    obj.loadloss(i)=obj.ld(i)-obj.wd(i)-obj.mx1*obj.unit1-obj.mx2*obj.unit2-obj.mx3*obj.unit3;
                else
                    obj.loadloss(i)=0;
                end%失负荷损失
            end
            obj.wdtotal=sum(obj.windloss);%总弃风
            obj.ldtotal=sum(obj.loadloss);%总失负荷
        end

        function f2(obj)
            for i=1:1:96
                if obj.wdtotal*0.9>=obj.ldtotal%弃风大于失负荷
                    rate=(obj.ldtotal/0.9)/obj.wdtotal;
                    if obj.windloss(i)>0
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store*(1-rate)*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        A=[-obj.unit1,-obj.unit2,-obj.unit3];
                        b=obj.wd(i)-obj.ld(i);
                        Aeq=[];
                        beq=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store*(1-rate)*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,5)=obj.store*(1-rate)*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%有弃风时
                    if obj.loadloss(i)>0
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+obj.store*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        A=[obj.unit1,obj.unit2,obj.unit3];
                        b=-obj.wd(i)+obj.ld(i);
                        Aeq=[];
                        beq=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+obj.store*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,6)=obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,8)=obj.store*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%有失负荷时
                    if obj.windloss(i)==obj.loadloss(i)
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        Aeq=[-obj.unit1,-obj.unit2,-obj.unit3];
                        beq=+obj.wd(i)-obj.ld(i);
                        A=[];
                        b=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%无弃风无失负荷时
                else%失负荷大于弃风
                    rate=obj.wdtotal/(obj.ldtotal/0.9);
                    if obj.windloss(i)>0
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store*0.9*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6+obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        A=[-obj.unit1,-obj.unit2,-obj.unit3];
                        b=obj.wd(i)-obj.ld(i);
                        Aeq=[];
                        beq=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store*0.9*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6+obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,5)=obj.store1*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*0.3*15*60/3.6;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,8)=obj.store*0.9*(-obj.ld(i)+(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*1.401*15*60/3.6;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%有弃风时
                    if obj.loadloss(i)>0
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+obj.store*(1-rate)*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        A=[obj.unit1,obj.unit2,obj.unit3];
                        b=-obj.wd(i)+obj.ld(i);
                        Aeq=[];
                        beq=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+obj.store*(1-rate)*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,6)=obj.store1*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6+obj.store*(1-rate)*(obj.ld(i)-(obj.wd(i)+x(1)*obj.unit1+x(2)*obj.unit2+x(3)*obj.unit3))*8*15*60/3.6;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%有失负荷时
                    if obj.windloss(i)==obj.loadloss(i)
                        fun= @(x)((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        Aeq=[-obj.unit1,-obj.unit2,-obj.unit3];
                        beq=+obj.wd(i)-obj.ld(i);
                        A=[];
                        b=[];
                        lb=[obj.mn1,obj.mn2,obj.mn3];
                        ub=[obj.mx1,obj.mx2,obj.mx3];
                        x0=[obj.mn1,obj.mn2,obj.mn3];
                        x=fmincon(fun,x0,A,b,Aeq,beq,lb,ub);
                        wealth=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5+0.045*obj.wd(i)*1000*0.25+(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,1:3)=[x(1)*obj.unit1,x(2)*obj.unit2,x(3)*obj.unit3];
                        obj.out(i,4)=wealth;
                        obj.out(i,7)=(obj.ce1*x(1)*obj.unit1+obj.ce2*x(2)*obj.unit2+obj.ce3*x(3)*obj.unit3)*obj.carbon*0.25;
                        obj.out(i,9)=0.045*obj.wd(i)*1000*0.25;
                        obj.out(i,10)=((obj.ca1*x(1)^2+obj.cb1*x(1)+obj.cc1)*obj.unit1+(obj.ca2*x(2)^2+obj.cb2*x(2)+obj.cc2)*obj.unit2+(obj.ca3*x(3)^2+obj.cb3*x(3)+obj.cc3)*obj.unit3)*0.7*0.25*1.5;
                    end%无弃风无失负荷时
                end
            end%out输出
        end

        function f3(obj)
            obj.machine1=obj.out(:,1);% 1:机组1功率
            obj.machine2=obj.out(:,2);% 2:机组2功率
            obj.machine3=obj.out(:,3);% 3:机组3功率
            obj.totalcost=sum(obj.out(:,4));% 4:总成本
            obj.abwindcost=sum(obj.out(:,5));% 5:弃风损失
            obj.loadcost=sum(obj.out(:,6));% 6:失负荷损失
            obj.carboncost=sum(obj.out(:,7));% 7:碳捕集成本
            obj.storecost=sum(obj.out(:,8));% 8:储能成本
            obj.windcost=sum(obj.out(:,9));% 9:风电成本
            obj.firecost=sum(obj.out(:,10));% 10:火电成本
            obj.unitprive=obj.totalcost/15507821;% 11:单价
        end
        
        function f4(obj)
            obj.f1();
            obj.f2();
            obj.f3();
        end
    end
end