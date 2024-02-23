# -*- coding: utf-8 -*
import time
from six.moves import input

import sys
import socket
import struct
import traceback
from math import pi
import ruamel.yaml
import os
import rospy
import xlrd, xlwt
from xlutils.copy import copy as xl_copy
import tf.transformations as tfm
import numpy as np
import keyboard
import re


from xyz_part_picking_vision.srv import PassiveProc



class PartPickerApp(object):
    def __init__(self):
        
        self._proxy = rospy.ServiceProxy('/part_picking_vision_server',
                                         PassiveProc)
        self._max_layer_num = 0
        self._return_result_mode_ = "single"
        self._grasp_poses = []
        self._vision_node_state = False

    def start_vision(self):
        try:
            self._proxy("0", "capture_images")
            self._vision_node_state = True
        except:
            print("wait vision ros server starting...")
            rospy.sleep(1)

    def run(self):
        while not rospy.is_shutdown():

            while True:
                if not self._vision_node_state:
                   # self.start_vision()
                    pass
                mode = 'P'
                tote_id = '0'
                input('a')


                filename1='test_data1.xls'
                path=os.path.join('/home/xyz/Pictures/python_excel/'+filename1)
                p1=os.path.exists(path)

                while p1:
                    flag1=raw_input("Excel is existed. Press d to Delete,Press c to Create NEW : ")
                    if flag1=='d':
                        os.remove(path)
                        break
                    if flag1=='c':
                        ret1=re.search('\d+',filename1)
                        count1=filename1.find(ret1.group(0))
                        filename1[count1]=ret1.group(0)+'1'
                        break
                    else :
                        print("error key")
                        continue
                wb=xlwt.Workbook(encoding='utf-8')
                sheet=wb.add_sheet('model1_0')
 
                #next
                # wb1=xlrd.open_workbook(filename1)
                # wb=xl_copy(wb1)
                # sheet=wb.add_sheet('model3_2')

                # 标题名
                title=['pos_x','pos_y','pos_z','eul_x','eul_y','eul_z']
                pose_data=[]
                pose_data_count=[]
                pose_data_new=[]
                result_sort_x=[]
                falg11=5

                # 写入表头
                i=0
                for j in title:
                    sheet.write(0,i,j)
                    i+=1
                #print (len(self._grasp_poses))
                
                if mode == 'P':
                    
                    #if len(self._grasp_poses) == 0:
                    cycle_times=int(input("input cycle times: "))
                    for f1 in range(cycle_times):
                        
                        try:
                            start = time.time()
                            response = self._proxy('0', 'capture_images', 0, [])
                            response = self._proxy('0','calculate_object_poses', 0, [])
                            end = time.time()
                            print("-------------------")
                        except:
                            self._vision_node_state = False
                            print("call vision servise failed")
                            
                        print(response) 
                        #print(response.objects.pose)
                        for obj in response.objects:
                            angles=tfm.euler_from_quaternion([obj.pose.orientation.x, obj.pose.orientation.y, obj.pose.orientation.z,obj.pose.orientation.w],axes='sxyz')
                            pose_data.append(obj.pose.position.x*1000)
                            pose_data.append(obj.pose.position.y*1000)
                            pose_data.append(obj.pose.position.z*1000)
                            pose_data.append(angles[0]*180/pi)
                            pose_data.append(angles[1]*180/pi)
                            pose_data.append(angles[2]*180/pi)
                        
                            print(pose_data)
                            #print(obj.pose.position.x)
                            pose_data_count.append(pose_data)
                            pose_data=[]
                            # pose_data_count=[]
                        result_sort_x=sorted(pose_data_count,key=lambda x:x[0],reverse= True)
                        pose_data_new.append(result_sort_x)
                        pose_data=[]
                        result_sort_x=[]
                        print("=======================================================================")
                        print(pose_data_new)
                        pose_data_count=[]

                    


                        # 写入表内容
                    # l=1
                    # for d in pose_data_new:
                        
                    #     c=0
                    #     for dd in d:
                    #         sheet.write(l,c,dd)
                    #         c+=1
                    #     l+=1

#-----------------------------------------------------------------------------------------------------------
                    # l=1
                    # for i in range(0,len(pose_data_new[0]),6):
                        
                    #     for j in range(cycle_times):
                    #         print pose_data_new[j][i:i+6]
                    #         c=0
                    #         for k in pose_data_new[j][i:i+6]:
                    #             print(k) 
                    #             sheet.write(l,c,k)
                    #             c+=1
                    #         l+=1
                    #     l+=1


                    l=1
                    for i in range(len(pose_data_new[0])):
                        for j in range(cycle_times):
                            c=0
                            for k in pose_data_new[j][i]:
                                sheet.write(l,c,k)
                                c+=1
                            l+=1
                        l+=1
                    print('Time:',end - start)
                    wb.save(filename1)
                    # # 保存

if __name__ == "__main__":

    app = PartPickerApp()
    app.run()
