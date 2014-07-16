# -*- coding: utf-8 -*-
"""
uiautil : UI Automation utility

Windows上でしか動きません。
"""
import comtypes
from comtypes import CoCreateInstance
import comtypes.client
from comtypes.gen.UIAutomationClient import *

__uia = None
__root_element = None

def __init():
    global __uia, __root_element
    __uia = CoCreateInstance(CUIAutomation._reg_clsid_,
                             interface=IUIAutomation,
                             clsctx=comtypes.CLSCTX_INPROC_SERVER)
    __root_element = __uia.GetRootElement()
    
def get_window_element(title):
    """titleに指定したウィンドウのAutomationElement を取得する
    デスクトップ上でtitleが重複している場合、最初に見つかったものを返す
    
    """
    win_element = __root_element.FindFirst(TreeScope_Children,
                                           __uia.CreatePropertyCondition(
                                           UIA_NamePropertyId, title))
    return win_element

def find_control(base_element, ctltype):
    """指定したbase elementのサブツリーから指定したコントロールタイプIDを持つエレメントのシーケンスを返す
    
    """
    condition = __uia.CreatePropertyCondition(UIA_ControlTypePropertyId, ctltype)
    ctl_elements = base_element.FindAll(TreeScope_Subtree, condition)
    return [ ctl_elements.GetElement(i) for i in range(ctl_elements.Length) ]
   
def lookup_by_name(elements, name):
    """指定したエレメントのシーケンスから、指定したNameプロパティを持つエレメントのうち、最初のものを返す。
    ヒットしなければ Noneを返す

   """
    for element in elements:
        if element.CurrentName == name:
            return element
    return None

def lookup_by_automationid(elements, id):
    """指定したエレメントのシーケンスから、指定したAutomationIdプロパティを持つエレメントのうち、最初のものを返す。
    ヒットしなければ Noneを返す

   """
    for element in elements:
        if element.CurrentAutomationId == id:
            return element
    return None
    
def click_button(element):
    """指定したelement をIUIAutomationInvokePattern.Invoke() でクリックする
    指定したelementのIsInvokePatternAvailableプロパティがFalseの場合何もしない
    
    """
    isClickable = element.GetCurrentPropertyValue(UIA_IsInvokePatternAvailablePropertyId)
    if isClickable == True:
        ptn = element.GetCurrentPattern(UIA_InvokePatternId)
        ptn.QueryInterface(IUIAutomationInvokePattern).Invoke()


if __name__ == '__main__':
    __init()
    uia = __uia
