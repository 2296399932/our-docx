import { ZERO } from '../../../../dataset/constant/Common'
import {
  AREA_CONTEXT_ATTR,
  EDITOR_ELEMENT_STYLE_ATTR,
  EDITOR_ROW_ATTR
} from '../../../../dataset/constant/Element'
import { ControlComponent } from '../../../../dataset/enum/Control'
import { ElementType } from '../../../../dataset/enum/Element'
import { IElement } from '../../../../interface/Element'
import { omitObject } from '../../../../utils'
import { formatElementContext } from '../../../../utils/element'
import { CanvasEvent } from '../../CanvasEvent'
import {  getUUID } from '../../../../utils'

export function enter(evt: KeyboardEvent, host: CanvasEvent) {
  console.log('enter键被按下');
  const draw = host.getDraw()
  if (draw.isReadonly()) return
  const rangeManager = draw.getRange()
  if (!rangeManager.getIsCanInput()) return
  const { startIndex, endIndex } = rangeManager.getRange()
  const isCollapsed = rangeManager.getIsCollapsed()
  const elementList = draw.getElementList()
  const startElement = elementList[startIndex]
  const endElement = elementList[endIndex]
  // 最后一个列表项行首回车取消列表设置
  if (
    isCollapsed &&
    endElement.listId &&
    endElement.value === ZERO &&
    elementList[endIndex + 1]?.listId !== endElement.listId
  ) {
    draw.getListParticle().unsetList()
    return
  }

  // 创建新段落元素
  let enterText: IElement = {
    value: ZERO,
    type: ElementType.PARAGRAPH,  // 设置为段落类型
    id: getUUID(),       // 生成唯一ID
    indent: draw.getOptions().defaultIndent || 0,  // 设置默认缩进
    valueList: []  // 初始化空的值列表
  }

  // 添加标记，表示这是通过Enter键创建的新段落
  ;(enterText as any).__isNewParagraph = true
  console.log('创建新段落元素', enterText);

  // 列表块内换行
  if (evt.shiftKey && startElement.listId) {
    enterText.listWrap = true
  }

  // 格式化上下文
  formatElementContext(elementList, [enterText], startIndex, {
    isBreakWhenWrap: true,
    editorOptions: draw.getOptions()
  })

  // shift长按 && 最后位置回车无需复制区域上下文
  if (
    evt.shiftKey &&
    endElement.areaId &&
    endElement.areaId !== elementList[endIndex + 1]?.areaId
  ) {
    enterText = omitObject(enterText, AREA_CONTEXT_ATTR)
  }

  // 标题结尾处回车无需格式化及样式复制
  if (
    !(
      endElement.id &&
      endElement.id !== elementList[endIndex + 1]?.id
    )
  ) {
    // 复制样式属性
    const copyElement = rangeManager.getRangeAnchorStyle(elementList, endIndex)
    if (copyElement) {
      const copyAttr = [...EDITOR_ROW_ATTR]
      // 不复制控件后缀样式
      if (copyElement.controlComponent !== ControlComponent.POSTFIX) {
        copyAttr.push(...EDITOR_ELEMENT_STYLE_ATTR)
      }
      copyAttr.forEach(attr => {
        const value = copyElement[attr] as never
        if (value !== undefined) {
          enterText[attr] = value
        }
      })
    }
  }

  // 控件或文档插入换行元素
  const control = draw.getControl()
  const activeControl = control.getActiveControl()
  let curIndex: number

  if (activeControl && control.getIsRangeWithinControl()) {
    curIndex = control.setValue([enterText])
    control.emitControlContentChange()
    console.log('在控件中插入新段落');
  } else {
    const position = draw.getPosition()
    const cursorPosition = position.getCursorPosition()
    if (!cursorPosition) return
    const { index } = cursorPosition

    if (isCollapsed) {
      console.log('在位置插入新段落', index + 1);
      draw.spliceElementList(elementList, index + 1, 0, [enterText])

      // 新增代码：更新后面文本的paragraphId
      const newId = enterText.id;
        console.log('新段落ID:', newId);

      // 遍历修改段落ID，直到遇到下一个段落标记或文档结束
      for (let i = index + 2; i < elementList.length; i++) {
        const element = elementList[i];

        // 如果遇到新的段落标记或零宽字符，停止修改 0.
        if (element.type === ElementType.PARAGRAPH ||element.type === ElementType.TITLE ||
          element.type === ElementType.TABLE ||
          element.type === ElementType.IMAGE ||
            (element.value === ZERO && element.type)) {
          console.log('遇到下一个段落标记，停止修改ID', i);
          break;
        }

        // 保存旧ID用于日志-
        const oldId = element.id;

        // 更新段落ID
        element.id = newId;
        console.log(`将元素 ${i} 的ID从 ${oldId} 改为 ${newId}`);

      }
    } else {
      console.log('替换选中内容并插入新段落', startIndex + 1);
      draw.spliceElementList(
        elementList,
        startIndex + 1,
        endIndex - startIndex,
        [enterText]
      )
    }
     console.log('修改后的元素', elementList)
    curIndex = index + 1
  }

  if (~curIndex) {
    rangeManager.setRange(curIndex, curIndex)
    console.log('设置光标位置并渲染', curIndex);
    draw.render({ curIndex })

    if (~curIndex && !isCollapsed) {
      // 只在光标在段落中间的情况下需要处理
      const newId = enterText.id;

      // 更新光标位置后的所有元素，直到遇到下一个段落标记
      for (let i = curIndex + 1; i < elementList.length; i++) {
        const element = elementList[i];

         // 如果遇到新的段落标记或其他块级元素，停止更新
  if (element.type === ElementType.PARAGRAPH ||
    element.type === ElementType.TITLE ||   // 添加标题类型
    element.type === ElementType.TABLE ||   // 添加表格类型
    element.type === ElementType.IMAGE ||   // 添加图片类型
    (element.value === ZERO && i !== curIndex)) {
  break;
}

        // 更新段落ID
        element.id = newId;
      }

      console.log('更新后面文本的段落ID为:', newId);
    }
  }

  evt.preventDefault()
}
