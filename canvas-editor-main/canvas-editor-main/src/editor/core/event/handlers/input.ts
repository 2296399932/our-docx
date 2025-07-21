import { ZERO } from '../../../dataset/constant/Common'
import {
  EDITOR_ELEMENT_COPY_ATTR,
  EDITOR_ELEMENT_STYLE_ATTR
} from '../../../dataset/constant/Element'
import { ElementType } from '../../../dataset/enum/Element'
import { IElement } from '../../../interface/Element'
import { IRangeElementStyle } from '../../../interface/Range'
import { splitText } from '../../../utils'
import { formatElementContext } from '../../../utils/element'
import { CanvasEvent } from '../../event/CanvasEvent'

export function input(data: string, host: CanvasEvent) {
  console.log('input函数被调用，输入数据:', data);
  const draw = host.getDraw()
  if (draw.isReadonly() || draw.isDisabled()) return
  const position = draw.getPosition()
  const cursorPosition = position.getCursorPosition()
  if (!data || !cursorPosition) return
  const isComposing = host.isComposing
  // 正在合成文本进行非输入操作
  if (isComposing && host.compositionInfo?.value === data) return
  const rangeManager = draw.getRange()
  if (!rangeManager.getIsCanInput()) return
  // 移除合成前，缓存设置的默认样式设置
  const defaultStyle =
    rangeManager.getDefaultStyle() || host.compositionInfo?.defaultStyle || null
  // 移除合成输入
  removeComposingInput(host)
  if (!isComposing) {
    const cursor = draw.getCursor()
    cursor.clearAgentDomValue()
  }
  const { TEXT, HYPERLINK, SUBSCRIPT, SUPERSCRIPT, DATE, TAB } = ElementType
  const text = data.replaceAll(`\n`, ZERO)
  const { startIndex, endIndex } = rangeManager.getRange()
  // 格式化元素
  const elementList = draw.getElementList()
  const copyElement = rangeManager.getRangeAnchorStyle(elementList, endIndex)
  if (!copyElement) return
  const isDesignMode = draw.isDesignMode()
  const inputData: IElement[] = splitText(text).map(value => {
    const newElement: IElement = {
      value
    }
    if (
      isDesignMode ||
      (!copyElement.title?.disabled && !copyElement.control?.disabled)
    ) {
      const nextElement = elementList[endIndex + 1]
      // 文本、超链接、日期、上下标：复制所有信息（元素类型、样式、特殊属性）
      if (
        !copyElement.type ||
        copyElement.type === TEXT ||
        (copyElement.type === HYPERLINK && nextElement?.type === HYPERLINK) ||
        (copyElement.type === DATE && nextElement?.type === DATE) ||
        (copyElement.type === SUBSCRIPT && nextElement?.type === SUBSCRIPT) ||
        (copyElement.type === SUPERSCRIPT && nextElement?.type === SUPERSCRIPT)
      ) {
        EDITOR_ELEMENT_COPY_ATTR.forEach(attr => {
          // 在分组外无需复制分组信息
          if (attr === 'groupIds' && !nextElement?.groupIds) return
          const value = copyElement[attr] as never
          if (value !== undefined) {
            newElement[attr] = value
          }
        })
      }
      // 仅复制样式：存在默认样式设置 || 无法匹配文本类元素时（TAB）
      if (defaultStyle || copyElement.type === TAB) {
        EDITOR_ELEMENT_STYLE_ATTR.forEach(attr => {
          const value =
            defaultStyle?.[attr as keyof IRangeElementStyle] ||
            copyElement[attr]
          if (value !== undefined) {
            newElement[attr] = value as never
          }
        })
      }
      if (isComposing) {
        newElement.underline = true
      }
    }
    return newElement
  })

  // 确保文本元素被包裹在段落中
  console.log('准备处理inputData，确保文本被包裹在段落中', inputData);
  const processedInputData = inputData;

  // 获取当前位置的元素，检查是否在段落内
  const currentElement = elementList[startIndex];
  console.log('当前位置元素:', currentElement);

  // 控件-移除placeholder
  const control = draw.getControl()
  let curIndex: number;

  if (control.getActiveControl() && control.getIsRangeWithinControl()) {
    curIndex = control.setValue(processedInputData)
    if (!isComposing) {
      control.emitControlContentChange()
    }
  } else {
    const start = startIndex + 1
    if (startIndex !== endIndex) {
      draw.spliceElementList(elementList, start, endIndex - startIndex)
    }
    formatElementContext(elementList, processedInputData, startIndex, {
      editorOptions: draw.getOptions()
    })

    // 如果当前在段落内，为文本元素添加段落ID
    if (currentElement && currentElement.type === ElementType.PARAGRAPH) {
      console.log('当前在段落内，使用另一种方式处理文本');

      // 为文本元素添加段落ID，使其与当前段落关联
      processedInputData.forEach(element => {
        element.id = currentElement.id;
      });

      console.log('插入关联到段落的元素:', processedInputData);
    } else if (currentElement && currentElement.id) {
      // Id（即在段落内的文本元素）
      console.log('当前在段落内的文本元素上，继承段落ID');

      // Id
      const id = currentElement.id;

      // 为文本元素添加相同的段落ID
      processedInputData.forEach(element => {
        element.id = id;
      });

      console.log('插入关联到段落的元素:', processedInputData);
    }

    // 插入处理后的元素
    draw.spliceElementList(elementList, start, 0, processedInputData)
    curIndex = startIndex + processedInputData.length
  }

  if (~curIndex) {
    rangeManager.setRange(curIndex, curIndex)
    draw.render({
      curIndex,
      isSubmitHistory: !isComposing
    })
  }
  if (isComposing) {
    host.compositionInfo = {
      elementList,
      value: text,
      startIndex: curIndex - processedInputData.length,
      endIndex: curIndex,
      defaultStyle
    }
  }
}

export function removeComposingInput(host: CanvasEvent) {
  if (!host.compositionInfo) return
  const { elementList, startIndex, endIndex } = host.compositionInfo
  elementList.splice(startIndex + 1, endIndex - startIndex)
  const rangeManager = host.getDraw().getRange()
  rangeManager.setRange(startIndex, startIndex)
  host.compositionInfo = null
}
