// import { commentList, data, options } from './mock'
import { IElement } from './editor'
const commentList: any[] = [];

const options: any = {};
import './style.css'
import prism from 'prismjs'

// API服务器地址配置
const API_BASE_URL = 'http://localhost:8000'

// 添加全局调试函数
declare global {
  interface Window {
    debugIndent: () => void;
    editor: any;
  }
}

import Editor, {
  BlockType,
  Command,
  ControlState,
  ControlType,
  EditorMode,
  EditorZone,
  ElementType,
  IBlock,
  ICatalogItem,
  KeyMap,
  ListStyle,
  ListType,
  PageMode,
  PaperDirection,
  RowFlex,
  TextDecorationStyle,
  TitleLevel,
  splitText
} from './editor'
import { Dialog } from './components/dialog/Dialog'
import { formatPrismToken } from './utils/prism'
import { Signature } from './components/signature/Signature'
import { debounce, nextTick, scrollIntoView } from './utils'

window.onload = function () {
  const isApple =
    typeof navigator !== 'undefined' && /Mac OS X/.test(navigator.userAgent)

  // 添加调试函数
  window.debugIndent = function() {
    const instance = window.editor;
    if (!instance) {
      console.error('编辑器实例不存在');
      return;
    }

    // 获取当前选中段落
    const elementList = instance.command.getRangeParagraph();
    if (!elementList || !elementList.length) {
      console.log('未选中段落元素');
      return;
    }


  };


  // 方法2：创建PARAGRAPH类型的元素

  // 1. 初始化编辑器
  const container = document.querySelector<HTMLDivElement>('.editor')!
  const instance = new Editor(
    container,
    {
      header: [

      ],
      main: [
        {
          type: ElementType.PARAGRAPH,
          value: "",
          paragraphId: "para-1",
          valueList: [
            { value: "这是一个" },
            { value: "复杂段落", bold: true },
            { value: "示例", italic: true },
            { value: "2", type: "subscript" }
          ]
        },
      {

        value: "O",
        font: "Microsoft YaHei",
        size: 16
      }],
      footer: [
        {
          value: 'canvas-editor',
          size: 12
        }
      ]
    },
    options
  )
  console.log('实例: ', instance)
  // cypress使用
  Reflect.set(window, 'editor', instance)

  // 修复所有缩进值
  instance.command.fixAllIndentValues();

  // 菜单弹窗销毁
  window.addEventListener(
    'click',
    evt => {
      const visibleDom = document.querySelector('.visible')
      if (!visibleDom || visibleDom.contains(<Node>evt.target)) return
      visibleDom.classList.remove('visible')
    },
    {
      capture: true
    }
  )

  // 2. | 撤销 | 重做 | 格式刷 | 清除格式 |
  const undoDom = document.querySelector<HTMLDivElement>('.menu-item__undo')!
  undoDom.title = `撤销(${isApple ? '⌘' : 'Ctrl'}+Z)`
  undoDom.onclick = function () {
    console.log('undo')
    instance.command.executeUndo()
  }

  const redoDom = document.querySelector<HTMLDivElement>('.menu-item__redo')!
  redoDom.title = `重做(${isApple ? '⌘' : 'Ctrl'}+Y)`
  redoDom.onclick = function () {
    console.log('redo')
    instance.command.executeRedo()
  }

  const painterDom = document.querySelector<HTMLDivElement>(
    '.menu-item__painter'
  )!
  const indentDom = document.querySelector<HTMLDivElement>('.menu-item__indent')
  if (indentDom) {
    const indentOptionDom = indentDom.querySelector<HTMLDivElement>('.options')!
    const indentValueInput = document.getElementById('indent-value') as HTMLInputElement
    const indentUnitSelect = document.getElementById('indent-unit') as HTMLSelectElement
    const indentApplyBtn = document.getElementById('indent-apply') as HTMLButtonElement
    const indentIncreaseBtn = document.getElementById('indent-increase') as HTMLButtonElement
    const indentDecreaseBtn = document.getElementById('indent-decrease') as HTMLButtonElement

    // 单位转换函数
    const toPt = (value: number, unit: string): number => {
      switch (unit) {
        case 'pt': return value;
        case 'cm': return value * 28.35; // 1厘米约等于28.35磅
        case 'ch': return value * 8; // 假设一个字符约为8磅
        default: return value;
      }
    }

    // 从磅转换到其他单位
    const fromPt = (value: number, unit: string): number => {
      switch (unit) {
        case 'pt': return value;
        case 'cm': return value / 28.35;
        case 'ch': return value / 8;
        default: return value;
      }
    }

    indentDom.onclick = function (e) {
      if (e.target === indentDom || e.target === indentDom.querySelector('i')) {
        indentOptionDom.classList.toggle('visible');

        // 获取当前选中段落的缩进值
        const elementList = instance.command.getRangeParagraph();
        if (elementList && elementList.length > 0) {
          const indent = elementList[0].indent || 0;
          indentValueInput.value = String(indent);
        }
      }
    }

    indentApplyBtn.onclick = function () {
      const value = parseInt(indentValueInput.value || '0');
      const unit = indentUnitSelect.value;
      const ptValue = toPt(value, unit);

      console.log('应用缩进值:', { value, unit, ptValue });

      // 获取选中段落元素
      const elementList = instance.command.getRangeParagraph();
      if (elementList && elementList.length > 0) {
        console.log('选中的段落元素:', JSON.stringify(elementList));

        // 直接使用command.indent和outdent方法更可靠
        // 先将缩进值重置为0
        let currentElements = elementList || [];
        let currentIndentValue = currentElements.length > 0 ? (currentElements[0].indent || 0) : 0;

        console.log('当前缩进值:', currentIndentValue);

        // 减少缩进直到为0
        let decreaseCount = 0;
        const maxIterations = 100; // 最大迭代次数，防止无限循环
        while (currentIndentValue > 0 && decreaseCount < maxIterations) {
          instance.command.executeOutdent();
          decreaseCount++;
          // 重新获取当前缩进值
          const updatedElements = instance.command.getRangeParagraph();
          currentElements = updatedElements || [];
          if (currentElements.length === 0) break;
          currentIndentValue = currentElements[0].indent || 0;
          console.log('减少缩进后的值:', currentIndentValue);
        }

        if (decreaseCount >= maxIterations) {
          console.warn('达到最大迭代次数限制，可能存在无限循环');
        }

        // 然后增加到所需的缩进值
        let targetIndent = Math.max(0, Math.min(ptValue, 100)); // 限制在0-100之间
        currentElements = instance.command.getRangeParagraph() || [];
        currentIndentValue = currentElements.length > 0 ? (currentElements[0].indent || 0) : 0;

        console.log('目标缩进值:', targetIndent, '当前缩进值:', currentIndentValue);

        // 增加缩进直到达到目标值
        let increaseCount = 0;
        while (currentIndentValue < targetIndent && increaseCount < maxIterations) {
          instance.command.executeIndent();
          increaseCount++;
          // 重新获取当前缩进值
          const updatedElements = instance.command.getRangeParagraph();
          currentElements = updatedElements || [];
          currentIndentValue = currentElements.length > 0 ? (currentElements[0].indent || 0) : 0;
          console.log('增加缩进后的值:', currentIndentValue);

          // 安全检查，避免无限循环
          if (currentIndentValue >= 100) break;
        }

        if (increaseCount >= maxIterations) {
          console.warn('达到最大迭代次数限制，可能存在无限循环');
        }

        // 强制更新渲染
        instance.command.executeForceUpdate({
          isSubmitHistory: true
        });

        // 调试输出
        console.log('缩进应用后的元素:', JSON.stringify(instance.command.getRangeParagraph()));
        console.log('当前文档值:', JSON.stringify(instance.command.getValue().data.main));

        // 使用我们添加的调试函数
        window.debugIndent();
      } else {
        console.log('未选中段落元素');
      }

      indentOptionDom.classList.remove('visible');
    }

    indentIncreaseBtn.onclick = function () {
      instance.command.executeIndent();
      indentOptionDom.classList.remove('visible');
    }

    indentDecreaseBtn.onclick = function () {
      instance.command.executeOutdent();
      indentOptionDom.classList.remove('visible');
    }

    // 单位变化时转换值
    indentUnitSelect.onchange = function () {
      const currentValue = parseInt(indentValueInput.value || '0');
      const oldUnit = indentUnitSelect.dataset.lastUnit || 'pt';
      const newUnit = indentUnitSelect.value;

      // 先转换为磅，再从磅转换为新单位
      const ptValue = toPt(currentValue, oldUnit);
      const newValue = fromPt(ptValue, newUnit);

      indentValueInput.value = Math.round(newValue * 100) / 100 + '';
      indentUnitSelect.dataset.lastUnit = newUnit;
    }
  }

  // 减少缩进按钮
  const outdentDom = document.querySelector<HTMLDivElement>('.menu-item__outdent')
  if (outdentDom) {
    outdentDom.onclick = function () {
      instance.command.executeOutdent()
    }
  }

  let isFirstClick = true
  let painterTimeout: number
  painterDom.onclick = function () {
    if (isFirstClick) {
      isFirstClick = false
      painterTimeout = window.setTimeout(() => {
        console.log('painter-click')
        isFirstClick = true
        instance.command.executePainter({
          isDblclick: false
        })
      }, 200)
    } else {
      window.clearTimeout(painterTimeout)
    }
  }

  painterDom.ondblclick = function () {
    console.log('painter-dblclick')
    isFirstClick = true
    window.clearTimeout(painterTimeout)
    instance.command.executePainter({
      isDblclick: true
    })
  }

  document.querySelector<HTMLDivElement>('.menu-item__format')!.onclick =
    function () {
      console.log('format')
      instance.command.executeFormat()
    }

  // 3. | 字体 | 字体变大 | 字体变小 | 加粗 | 斜体 | 下划线 | 删除线 | 上标 | 下标 | 字体颜色 | 背景色 |
  const fontDom = document.querySelector<HTMLDivElement>('.menu-item__font')!
  const fontSelectDom = fontDom.querySelector<HTMLDivElement>('.select')!
  const fontOptionDom = fontDom.querySelector<HTMLDivElement>('.options')!
  fontDom.onclick = function () {
    console.log('font')
    fontOptionDom.classList.toggle('visible')
  }
  fontOptionDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    instance.command.executeFont(li.dataset.family!)
  }

  const sizeSetDom = document.querySelector<HTMLDivElement>('.menu-item__size')!
  const sizeSelectDom = sizeSetDom.querySelector<HTMLDivElement>('.select')!
  const sizeOptionDom = sizeSetDom.querySelector<HTMLDivElement>('.options')!
  sizeSetDom.title = `设置字号`
  sizeSetDom.onclick = function () {
    console.log('size')
    sizeOptionDom.classList.toggle('visible')
  }
  sizeOptionDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    instance.command.executeSize(Number(li.dataset.size!))
  }

  const sizeAddDom = document.querySelector<HTMLDivElement>(
    '.menu-item__size-add'
  )!
  sizeAddDom.title = `增大字号(${isApple ? '⌘' : 'Ctrl'}+[)`
  sizeAddDom.onclick = function () {
    console.log('size-add')
    instance.command.executeSizeAdd()
  }

  const sizeMinusDom = document.querySelector<HTMLDivElement>(
    '.menu-item__size-minus'
  )!
  sizeMinusDom.title = `减小字号(${isApple ? '⌘' : 'Ctrl'}+])`
  sizeMinusDom.onclick = function () {
    console.log('size-minus')
    instance.command.executeSizeMinus()
  }

  const boldDom = document.querySelector<HTMLDivElement>('.menu-item__bold')!
  boldDom.title = `加粗(${isApple ? '⌘' : 'Ctrl'}+B)`
  boldDom.onclick = function () {
    console.log('bold')
    instance.command.executeBold()
  }

  const italicDom =
    document.querySelector<HTMLDivElement>('.menu-item__italic')!
  italicDom.title = `斜体(${isApple ? '⌘' : 'Ctrl'}+I)`
  italicDom.onclick = function () {
    console.log('italic')
    instance.command.executeItalic()
  }

  const underlineDom = document.querySelector<HTMLDivElement>(
    '.menu-item__underline'
  )!
  underlineDom.title = `下划线(${isApple ? '⌘' : 'Ctrl'}+U)`
  const underlineOptionDom =
    underlineDom.querySelector<HTMLDivElement>('.options')!
  underlineDom.querySelector<HTMLSpanElement>('.select')!.onclick =
    function () {
      underlineOptionDom.classList.toggle('visible')
    }
  underlineDom.querySelector<HTMLElement>('i')!.onclick = function () {
    console.log('underline')
    instance.command.executeUnderline()
    underlineOptionDom.classList.remove('visible')
  }
  underlineDom.querySelector<HTMLUListElement>('ul')!.onmousedown = function (
    evt
  ) {
    const li = evt.target as HTMLLIElement
    const decorationStyle = <TextDecorationStyle>li.dataset.decorationStyle
    instance.command.executeUnderline({
      style: decorationStyle
    })
    underlineOptionDom.classList.remove('visible')
  }

  const strikeoutDom = document.querySelector<HTMLDivElement>(
    '.menu-item__strikeout'
  )!
  strikeoutDom.onclick = function () {
    console.log('strikeout')
    instance.command.executeStrikeout()
  }

  const superscriptDom = document.querySelector<HTMLDivElement>(
    '.menu-item__superscript'
  )!
  superscriptDom.title = `上标(${isApple ? '⌘' : 'Ctrl'}+Shift+,)`
  superscriptDom.onclick = function () {
    console.log('superscript')
    instance.command.executeSuperscript()
  }

  const subscriptDom = document.querySelector<HTMLDivElement>(
    '.menu-item__subscript'
  )!
  subscriptDom.title = `下标(${isApple ? '⌘' : 'Ctrl'}+Shift+.)`
  subscriptDom.onclick = function () {
    console.log('subscript')
    instance.command.executeSubscript()
  }

  const colorControlDom = document.querySelector<HTMLInputElement>('#color')!
  colorControlDom.oninput = function () {
    instance.command.executeColor(colorControlDom.value)
  }
  const colorDom = document.querySelector<HTMLDivElement>('.menu-item__color')!
  const colorSpanDom = colorDom.querySelector('span')!
  colorDom.onclick = function () {
    console.log('color')
    colorControlDom.click()
  }

  const highlightControlDom =
    document.querySelector<HTMLInputElement>('#highlight')!
  highlightControlDom.oninput = function () {
    instance.command.executeHighlight(highlightControlDom.value)
  }
  const highlightDom = document.querySelector<HTMLDivElement>(
    '.menu-item__highlight'
  )!
  const highlightSpanDom = highlightDom.querySelector('span')!
  highlightDom.onclick = function () {
    console.log('highlight')
    highlightControlDom?.click()
  }

  const titleDom = document.querySelector<HTMLDivElement>('.menu-item__title')!
  const titleSelectDom = titleDom.querySelector<HTMLDivElement>('.select')!
  const titleOptionDom = titleDom.querySelector<HTMLDivElement>('.options')!
  titleOptionDom.querySelectorAll('li').forEach((li, index) => {
    li.title = `Ctrl+${isApple ? 'Option' : 'Alt'}+${index}`
  })

  titleDom.onclick = function () {
    console.log('title')
    titleOptionDom.classList.toggle('visible')
  }
  titleOptionDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    const level = <TitleLevel>li.dataset.level
    instance.command.executeTitle(level || null)
  }

  const leftDom = document.querySelector<HTMLDivElement>('.menu-item__left')!
  leftDom.title = `左对齐(${isApple ? '⌘' : 'Ctrl'}+L)`
  leftDom.onclick = function () {
    console.log('left')
    instance.command.executeRowFlex(RowFlex.LEFT)
  }

  const centerDom =
    document.querySelector<HTMLDivElement>('.menu-item__center')!
  centerDom.title = `居中对齐(${isApple ? '⌘' : 'Ctrl'}+E)`
  centerDom.onclick = function () {
    console.log('center')
    instance.command.executeRowFlex(RowFlex.CENTER)
  }

  const rightDom = document.querySelector<HTMLDivElement>('.menu-item__right')!
  rightDom.title = `右对齐(${isApple ? '⌘' : 'Ctrl'}+R)`
  rightDom.onclick = function () {
    console.log('right')
    instance.command.executeRowFlex(RowFlex.RIGHT)
  }

  const alignmentDom = document.querySelector<HTMLDivElement>(
    '.menu-item__alignment'
  )!
  alignmentDom.title = `两端对齐(${isApple ? '⌘' : 'Ctrl'}+J)`
  alignmentDom.onclick = function () {
    console.log('alignment')
    instance.command.executeRowFlex(RowFlex.ALIGNMENT)
  }

  const justifyDom = document.querySelector<HTMLDivElement>(
    '.menu-item__justify'
  )!
  justifyDom.title = `分散对齐(${isApple ? '⌘' : 'Ctrl'}+Shift+J)`
  justifyDom.onclick = function () {
    console.log('justify')
    instance.command.executeRowFlex(RowFlex.JUSTIFY)
  }

  const rowMarginDom = document.querySelector<HTMLDivElement>(
    '.menu-item__row-margin'
  )!
  const rowOptionDom = rowMarginDom.querySelector<HTMLDivElement>('.options')!
  rowMarginDom.onclick = function () {
    console.log('row-margin')
    rowOptionDom.classList.toggle('visible')
  }
  rowOptionDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    instance.command.executeRowMargin(Number(li.dataset.rowmargin!))
  }

  const listDom = document.querySelector<HTMLDivElement>('.menu-item__list')!
  listDom.title = `列表(${isApple ? '⌘' : 'Ctrl'}+Shift+U)`
  const listOptionDom = listDom.querySelector<HTMLDivElement>('.options')!
  listDom.onclick = function () {
    console.log('list')
    listOptionDom.classList.toggle('visible')
  }
  listOptionDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    const listType = <ListType>li.dataset.listType || null
    const listStyle = <ListStyle>(<unknown>li.dataset.listStyle)
    instance.command.executeList(listType, listStyle)
  }

  // 4. | 表格 | 图片 | 超链接 | 分割线 | 水印 | 代码块 | 分隔符 | 控件 | 复选框 | LaTeX | 日期选择器
  const tableDom = document.querySelector<HTMLDivElement>('.menu-item__table')!
  const tablePanelContainer = document.querySelector<HTMLDivElement>(
    '.menu-item__table__collapse'
  )!
  const tableClose = document.querySelector<HTMLDivElement>('.table-close')!
  const tableTitle = document.querySelector<HTMLDivElement>('.table-select')!
  const tablePanel = document.querySelector<HTMLDivElement>('.table-panel')!
  // 绘制行列
  const tableCellList: HTMLDivElement[][] = []
  for (let i = 0; i < 10; i++) {
    const tr = document.createElement('tr')
    tr.classList.add('table-row')
    const trCellList: HTMLDivElement[] = []
    for (let j = 0; j < 10; j++) {
      const td = document.createElement('td')
      td.classList.add('table-cel')
      tr.append(td)
      trCellList.push(td)
    }
    tablePanel.append(tr)
    tableCellList.push(trCellList)
  }
  let colIndex = 0
  let rowIndex = 0
  // 移除所有格选择
  function removeAllTableCellSelect() {
    tableCellList.forEach(tr => {
      tr.forEach(td => td.classList.remove('active'))
    })
  }
  // 设置标题内容
  function setTableTitle(payload: string) {
    tableTitle.innerText = payload
  }
  // 恢复初始状态
  function recoveryTable() {
    // 还原选择样式、标题、选择行列
    removeAllTableCellSelect()
    setTableTitle('插入')
    colIndex = 0
    rowIndex = 0
    // 隐藏panel
    tablePanelContainer.style.display = 'none'
  }
  tableDom.onclick = function () {
    console.log('table')
    tablePanelContainer!.style.display = 'block'
  }
  tablePanel.onmousemove = function (evt) {
    const celSize = 16
    const rowMarginTop = 10
    const celMarginRight = 6
    const { offsetX, offsetY } = evt
    // 移除所有选择
    removeAllTableCellSelect()
    colIndex = Math.ceil(offsetX / (celSize + celMarginRight)) || 1
    rowIndex = Math.ceil(offsetY / (celSize + rowMarginTop)) || 1
    // 改变选择样式
    tableCellList.forEach((tr, trIndex) => {
      tr.forEach((td, tdIndex) => {
        if (tdIndex < colIndex && trIndex < rowIndex) {
          td.classList.add('active')
        }
      })
    })
    // 改变表格标题
    setTableTitle(`${rowIndex}×${colIndex}`)
  }
  tableClose.onclick = function () {
    recoveryTable()
  }
  tablePanel.onclick = function () {
    // 应用选择
    instance.command.executeInsertTable(rowIndex, colIndex)
    recoveryTable()
  }

  const imageDom = document.querySelector<HTMLDivElement>('.menu-item__image')!
  const imageFileDom = document.querySelector<HTMLInputElement>('#image')!
  imageDom.onclick = function () {
    imageFileDom.click()
  }
  imageFileDom.onchange = function () {
    const file = imageFileDom.files![0]!
    const fileReader = new FileReader()
    fileReader.readAsDataURL(file)
    fileReader.onload = function () {
      // 计算宽高
      const image = new Image()
      const value = fileReader.result as string
      image.src = value
      image.onload = function () {
        instance.command.executeImage({
          value,
          width: image.width,
          height: image.height
        })
        imageFileDom.value = ''
      }
    }
  }

  const hyperlinkDom = document.querySelector<HTMLDivElement>(
    '.menu-item__hyperlink'
  )!
  hyperlinkDom.onclick = function () {
    console.log('hyperlink')
    new Dialog({
      title: '超链接',
      data: [
        {
          type: 'text',
          label: '文本',
          name: 'name',
          required: true,
          placeholder: '请输入文本',
          value: instance.command.getRangeText()
        },
        {
          type: 'text',
          label: '链接',
          name: 'url',
          required: true,
          placeholder: '请输入链接'
        }
      ],
      onConfirm: payload => {
        const name = payload.find(p => p.name === 'name')?.value
        if (!name) return
        const url = payload.find(p => p.name === 'url')?.value
        if (!url) return
        instance.command.executeHyperlink({
          type: ElementType.HYPERLINK,
          value: '',
          url,
          valueList: splitText(name).map(n => ({
            value: n,
            size: 16
          }))
        })
      }
    })
  }

  const separatorDom = document.querySelector<HTMLDivElement>(
    '.menu-item__separator'
  )!
  const separatorOptionDom =
    separatorDom.querySelector<HTMLDivElement>('.options')!
  separatorDom.onclick = function () {
    console.log('separator')
    separatorOptionDom.classList.toggle('visible')
  }
  separatorOptionDom.onmousedown = function (evt) {
    let payload: number[] = []
    const li = evt.target as HTMLLIElement
    const separatorDash = li.dataset.separator?.split(',').map(Number)
    if (separatorDash) {
      const isSingleLine = separatorDash.every(d => d === 0)
      if (!isSingleLine) {
        payload = separatorDash
      }
    }
    instance.command.executeSeparator(payload)
  }

  const pageBreakDom = document.querySelector<HTMLDivElement>(
    '.menu-item__page-break'
  )!
  pageBreakDom.onclick = function () {
    console.log('pageBreak')
    instance.command.executePageBreak()
  }

  const watermarkDom = document.querySelector<HTMLDivElement>(
    '.menu-item__watermark'
  )!
  const watermarkOptionDom =
    watermarkDom.querySelector<HTMLDivElement>('.options')!
  watermarkDom.onclick = function () {
    console.log('watermark')
    watermarkOptionDom.classList.toggle('visible')
  }
  watermarkOptionDom.onmousedown = function (evt) {
    const li = evt.target as HTMLLIElement
    const menu = li.dataset.menu!
    watermarkOptionDom.classList.toggle('visible')
    if (menu === 'add') {
      new Dialog({
        title: '水印',
        data: [
          {
            type: 'text',
            label: '内容',
            name: 'data',
            required: true,
            placeholder: '请输入内容'
          },
          {
            type: 'color',
            label: '颜色',
            name: 'color',
            required: true,
            value: '#AEB5C0'
          },
          {
            type: 'number',
            label: '字体大小',
            name: 'size',
            required: true,
            value: '120'
          },
          {
            type: 'number',
            label: '透明度',
            name: 'opacity',
            required: true,
            value: '0.3'
          },
          {
            type: 'select',
            label: '重复',
            name: 'repeat',
            value: '0',
            required: false,
            options: [
              {
                label: '不重复',
                value: '0'
              },
              {
                label: '重复',
                value: '1'
              }
            ]
          },
          {
            type: 'number',
            label: '水平间隔',
            name: 'horizontalGap',
            required: false,
            value: '10'
          },
          {
            type: 'number',
            label: '垂直间隔',
            name: 'verticalGap',
            required: false,
            value: '10'
          }
        ],
        onConfirm: payload => {
          const nullableIndex = payload.findIndex(p => !p.value)
          if (~nullableIndex) return
          const watermark = payload.reduce((pre, cur) => {
            pre[cur.name] = cur.value
            return pre
          }, <any>{})
          const repeat = watermark.repeat === '1'
          instance.command.executeAddWatermark({
            data: watermark.data,
            color: watermark.color,
            size: Number(watermark.size),
            opacity: Number(watermark.opacity),
            repeat,
            gap:
              repeat && watermark.horizontalGap && watermark.verticalGap
                ? [
                    Number(watermark.horizontalGap),
                    Number(watermark.verticalGap)
                  ]
                : undefined
          })
        }
      })
    } else {
      instance.command.executeDeleteWatermark()
    }
  }

  const codeblockDom = document.querySelector<HTMLDivElement>(
    '.menu-item__codeblock'
  )!
  codeblockDom.onclick = function () {
    console.log('codeblock')
    new Dialog({
      title: '代码块',
      data: [
        {
          type: 'textarea',
          name: 'codeblock',
          placeholder: '请输入代码',
          width: 500,
          height: 300
        }
      ],
      onConfirm: payload => {
        const codeblock = payload.find(p => p.name === 'codeblock')?.value
        if (!codeblock) return
        const tokenList = prism.tokenize(codeblock, prism.languages.javascript)
        const formatTokenList = formatPrismToken(tokenList)
        const elementList: IElement[] = []
        for (let i = 0; i < formatTokenList.length; i++) {
          const formatToken = formatTokenList[i]
          const tokenStringList = splitText(formatToken.content)
          for (let j = 0; j < tokenStringList.length; j++) {
            const value = tokenStringList[j]
            const element: IElement = {
              value
            }
            if (formatToken.color) {
              element.color = formatToken.color
            }
            if (formatToken.bold) {
              element.bold = true
            }
            if (formatToken.italic) {
              element.italic = true
            }
            elementList.push(element)
          }
        }
        elementList.unshift({
          value: '\n'
        })
        instance.command.executeInsertElementList(elementList)
      }
    })
  }

  const controlDom = document.querySelector<HTMLDivElement>(
    '.menu-item__control'
  )!
  const controlOptionDom = controlDom.querySelector<HTMLDivElement>('.options')!
  controlDom.onclick = function () {
    console.log('control')
    controlOptionDom.classList.toggle('visible')
  }
  controlOptionDom.onmousedown = function (evt) {
    controlOptionDom.classList.toggle('visible')
    const li = evt.target as HTMLLIElement
    const type = <ControlType>li.dataset.control
    switch (type) {
      case ControlType.TEXT:
        new Dialog({
          title: '文本控件',
          data: [
            {
              type: 'text',
              label: '占位符',
              name: 'placeholder',
              required: true,
              placeholder: '请输入占位符'
            },
            {
              type: 'text',
              label: '默认值',
              name: 'value',
              placeholder: '请输入默认值'
            }
          ],
          onConfirm: payload => {
            const placeholder = payload.find(
              p => p.name === 'placeholder'
            )?.value
            if (!placeholder) return
            const value = payload.find(p => p.name === 'value')?.value || ''
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                value: value
                  ? [
                      {
                        value
                      }
                    ]
                  : null,
                placeholder
              }
            })
          }
        })
        break
      case ControlType.SELECT:
        new Dialog({
          title: '列举控件',
          data: [
            {
              type: 'text',
              label: '占位符',
              name: 'placeholder',
              required: true,
              placeholder: '请输入占位符'
            },
            {
              type: 'text',
              label: '默认值',
              name: 'code',
              placeholder: '请输入默认值'
            },
            {
              type: 'textarea',
              label: '值集',
              name: 'valueSets',
              required: true,
              height: 100,
              placeholder: `请输入值集JSON，例：\n[{\n"value":"有",\n"code":"98175"\n}]`
            }
          ],
          onConfirm: payload => {
            const placeholder = payload.find(
              p => p.name === 'placeholder'
            )?.value
            if (!placeholder) return
            const valueSets = payload.find(p => p.name === 'valueSets')?.value
            if (!valueSets) return
            const code = payload.find(p => p.name === 'code')?.value
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                code,
                value: null,
                placeholder,
                valueSets: JSON.parse(valueSets)
              }
            })
          }
        })
        break
      case ControlType.CHECKBOX:
        new Dialog({
          title: '复选框控件',
          data: [
            {
              type: 'text',
              label: '默认值',
              name: 'code',
              placeholder: '请输入默认值，多个值以英文逗号分割'
            },
            {
              type: 'textarea',
              label: '值集',
              name: 'valueSets',
              required: true,
              height: 100,
              placeholder: `请输入值集JSON，例：\n[{\n"value":"有",\n"code":"98175"\n}]`
            }
          ],
          onConfirm: payload => {
            const valueSets = payload.find(p => p.name === 'valueSets')?.value
            if (!valueSets) return
            const code = payload.find(p => p.name === 'code')?.value
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                code,
                value: null,
                valueSets: JSON.parse(valueSets)
              }
            })
          }
        })
        break
      case ControlType.RADIO:
        new Dialog({
          title: '单选框控件',
          data: [
            {
              type: 'text',
              label: '默认值',
              name: 'code',
              placeholder: '请输入默认值'
            },
            {
              type: 'textarea',
              label: '值集',
              name: 'valueSets',
              required: true,
              height: 100,
              placeholder: `请输入值集JSON，例：\n[{\n"value":"有",\n"code":"98175"\n}]`
            }
          ],
          onConfirm: payload => {
            const valueSets = payload.find(p => p.name === 'valueSets')?.value
            if (!valueSets) return
            const code = payload.find(p => p.name === 'code')?.value
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                code,
                value: null,
                valueSets: JSON.parse(valueSets)
              }
            })
          }
        })
        break
      case ControlType.DATE:
        new Dialog({
          title: '日期控件',
          data: [
            {
              type: 'text',
              label: '占位符',
              name: 'placeholder',
              required: true,
              placeholder: '请输入占位符'
            },
            {
              type: 'text',
              label: '默认值',
              name: 'value',
              placeholder: '请输入默认值'
            },
            {
              type: 'select',
              label: '日期格式',
              name: 'dateFormat',
              value: 'yyyy-MM-dd hh:mm:ss',
              required: true,
              options: [
                {
                  label: 'yyyy-MM-dd hh:mm:ss',
                  value: 'yyyy-MM-dd hh:mm:ss'
                },
                {
                  label: 'yyyy-MM-dd',
                  value: 'yyyy-MM-dd'
                }
              ]
            }
          ],
          onConfirm: payload => {
            const placeholder = payload.find(
              p => p.name === 'placeholder'
            )?.value
            if (!placeholder) return
            const value = payload.find(p => p.name === 'value')?.value || ''
            const dateFormat =
              payload.find(p => p.name === 'dateFormat')?.value || ''
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                dateFormat,
                value: value
                  ? [
                      {
                        value
                      }
                    ]
                  : null,
                placeholder
              }
            })
          }
        })
        break
      case ControlType.NUMBER:
        new Dialog({
          title: '数值控件',
          data: [
            {
              type: 'text',
              label: '占位符',
              name: 'placeholder',
              required: true,
              placeholder: '请输入占位符'
            },
            {
              type: 'text',
              label: '默认值',
              name: 'value',
              placeholder: '请输入默认值'
            }
          ],
          onConfirm: payload => {
            const placeholder = payload.find(
              p => p.name === 'placeholder'
            )?.value
            if (!placeholder) return
            const value = payload.find(p => p.name === 'value')?.value || ''
            instance.command.executeInsertControl({
              type: ElementType.CONTROL,
              value: '',
              control: {
                type,
                value: value
                  ? [
                      {
                        value
                      }
                    ]
                  : null,
                placeholder
              }
            })
          }
        })
        break
      default:
        break
    }
  }

  const checkboxDom = document.querySelector<HTMLDivElement>(
    '.menu-item__checkbox'
  )!
  checkboxDom.onclick = function () {
    console.log('checkbox')
    instance.command.executeInsertElementList([
      {
        type: ElementType.CHECKBOX,
        checkbox: {
          value: false
        },
        value: ''
      }
    ])
  }

  const radioDom = document.querySelector<HTMLDivElement>('.menu-item__radio')!
  radioDom.onclick = function () {
    console.log('radio')
    instance.command.executeInsertElementList([
      {
        type: ElementType.RADIO,
        checkbox: {
          value: false
        },
        value: ''
      }
    ])
  }

  const latexDom = document.querySelector<HTMLDivElement>('.menu-item__latex')!
  latexDom.onclick = function () {
    console.log('LaTeX')
    new Dialog({
      title: 'LaTeX',
      data: [
        {
          type: 'textarea',
          height: 100,
          name: 'value',
          placeholder: '请输入LaTeX文本'
        }
      ],
      onConfirm: payload => {
        const value = payload.find(p => p.name === 'value')?.value
        if (!value) return
        instance.command.executeInsertElementList([
          {
            type: ElementType.LATEX,
            value
          }
        ])
      }
    })
  }

  const dateDom = document.querySelector<HTMLDivElement>('.menu-item__date')!
  const dateDomOptionDom = dateDom.querySelector<HTMLDivElement>('.options')!
  dateDom.onclick = function () {
    console.log('date')
    dateDomOptionDom.classList.toggle('visible')
    // 定位调整
    const bodyRect = document.body.getBoundingClientRect()
    const dateDomOptionRect = dateDomOptionDom.getBoundingClientRect()
    if (dateDomOptionRect.left + dateDomOptionRect.width > bodyRect.width) {
      dateDomOptionDom.style.right = '0px'
      dateDomOptionDom.style.left = 'unset'
    } else {
      dateDomOptionDom.style.right = 'unset'
      dateDomOptionDom.style.left = '0px'
    }
    // 当前日期
    const date = new Date()
    const year = date.getFullYear().toString()
    const month = (date.getMonth() + 1).toString().padStart(2, '0')
    const day = date.getDate().toString().padStart(2, '0')
    const hour = date.getHours().toString().padStart(2, '0')
    const minute = date.getMinutes().toString().padStart(2, '0')
    const second = date.getSeconds().toString().padStart(2, '0')
    const dateString = `${year}-${month}-${day}`
    const dateTimeString = `${dateString} ${hour}:${minute}:${second}`
    dateDomOptionDom.querySelector<HTMLLIElement>('li:first-child')!.innerText =
      dateString
    dateDomOptionDom.querySelector<HTMLLIElement>('li:last-child')!.innerText =
      dateTimeString
  }
  dateDomOptionDom.onmousedown = function (evt) {
    const li = evt.target as HTMLLIElement
    const dateFormat = li.dataset.format!
    dateDomOptionDom.classList.toggle('visible')
    instance.command.executeInsertElementList([
      {
        type: ElementType.DATE,
        value: '',
        dateFormat,
        valueList: [
          {
            value: li.innerText.trim()
          }
        ]
      }
    ])
  }

  const blockDom = document.querySelector<HTMLDivElement>('.menu-item__block')!
  blockDom.onclick = function () {
    console.log('block')
    new Dialog({
      title: '内容块',
      data: [
        {
          type: 'select',
          label: '类型',
          name: 'type',
          value: 'iframe',
          required: true,
          options: [
            {
              label: '网址',
              value: 'iframe'
            },
            {
              label: '视频',
              value: 'video'
            }
          ]
        },
        {
          type: 'number',
          label: '宽度',
          name: 'width',
          placeholder: '请输入宽度（默认页面内宽度）'
        },
        {
          type: 'number',
          label: '高度',
          name: 'height',
          required: true,
          placeholder: '请输入高度'
        },
        {
          type: 'input',
          label: '地址',
          name: 'src',
          required: false,
          placeholder: '请输入地址'
        },
        {
          type: 'textarea',
          label: 'HTML',
          height: 100,
          name: 'srcdoc',
          required: false,
          placeholder: '请输入HTML代码（仅网址类型有效）'
        }
      ],
      onConfirm: payload => {
        const type = payload.find(p => p.name === 'type')?.value
        if (!type) return
        const width = payload.find(p => p.name === 'width')?.value
        const height = payload.find(p => p.name === 'height')?.value
        if (!height) return
        // 地址或HTML代码至少存在一项
        const src = payload.find(p => p.name === 'src')?.value
        const srcdoc = payload.find(p => p.name === 'srcdoc')?.value
        const block: IBlock = {
          type: <BlockType>type
        }
        if (block.type === BlockType.IFRAME) {
          if (!src && !srcdoc) return
          block.iframeBlock = {
            src,
            srcdoc
          }
        } else if (block.type === BlockType.VIDEO) {
          if (!src) return
          block.videoBlock = {
            src
          }
        }
        const blockElement: IElement = {
          type: ElementType.BLOCK,
          value: '',
          height: Number(height),
          block
        }
        if (width) {
          blockElement.width = Number(width)
        }
        instance.command.executeInsertElementList([blockElement])
      }
    })
  }

  // 5. | 搜索&替换 | 打印 |
  const searchCollapseDom = document.querySelector<HTMLDivElement>(
    '.menu-item__search__collapse'
  )!
  const searchInputDom = document.querySelector<HTMLInputElement>(
    '.menu-item__search__collapse__search input'
  )!
  const replaceInputDom = document.querySelector<HTMLInputElement>(
    '.menu-item__search__collapse__replace input'
  )!
  const searchDom =
    document.querySelector<HTMLDivElement>('.menu-item__search')!
  searchDom.title = `搜索与替换(${isApple ? '⌘' : 'Ctrl'}+F)`
  const searchResultDom =
    searchCollapseDom.querySelector<HTMLLabelElement>('.search-result')!
  function setSearchResult() {
    const result = instance.command.getSearchNavigateInfo()
    if (result) {
      const { index, count } = result
      searchResultDom.innerText = `${index}/${count}`
    } else {
      searchResultDom.innerText = ''
    }
  }
  searchDom.onclick = function () {
    console.log('search')
    searchCollapseDom.style.display = 'block'
    const bodyRect = document.body.getBoundingClientRect()
    const searchRect = searchDom.getBoundingClientRect()
    const searchCollapseRect = searchCollapseDom.getBoundingClientRect()
    if (searchRect.left + searchCollapseRect.width > bodyRect.width) {
      searchCollapseDom.style.right = '0px'
      searchCollapseDom.style.left = 'unset'
    } else {
      searchCollapseDom.style.right = 'unset'
    }
    searchInputDom.focus()
  }
  searchCollapseDom.querySelector<HTMLSpanElement>('span')!.onclick =
    function () {
      searchCollapseDom.style.display = 'none'
      searchInputDom.value = ''
      replaceInputDom.value = ''
      instance.command.executeSearch(null)
      setSearchResult()
    }
  searchInputDom.oninput = function () {
    instance.command.executeSearch(searchInputDom.value || null)
    setSearchResult()
  }
  searchInputDom.onkeydown = function (evt) {
    if (evt.key === 'Enter') {
      instance.command.executeSearch(searchInputDom.value || null)
      setSearchResult()
    }
  }
  searchCollapseDom.querySelector<HTMLButtonElement>('button')!.onclick =
    function () {
      const searchValue = searchInputDom.value
      const replaceValue = replaceInputDom.value
      if (searchValue && replaceValue && searchValue !== replaceValue) {
        instance.command.executeReplace(replaceValue)
      }
    }
  searchCollapseDom.querySelector<HTMLDivElement>('.arrow-left')!.onclick =
    function () {
      instance.command.executeSearchNavigatePre()
      setSearchResult()
    }
  searchCollapseDom.querySelector<HTMLDivElement>('.arrow-right')!.onclick =
    function () {
      instance.command.executeSearchNavigateNext()
      setSearchResult()
    }

  const printDom = document.querySelector<HTMLDivElement>('.menu-item__print')!
  printDom.title = `打印(${isApple ? '⌘' : 'Ctrl'}+P)`
  printDom.onclick = function () {
    console.log('print')
    instance.command.executePrint()
  }

  // 上传文档功能
  const uploadDom = document.querySelector<HTMLDivElement>('.menu-item__upload')!
  const uploadFileDom = document.querySelector<HTMLInputElement>('#upload')!
  uploadDom.onclick = function () {
    uploadFileDom.click()
  }
  uploadFileDom.onchange = function () {
    const file = uploadFileDom.files![0]
    if (!file) return

    // 显示加载提示
    const loadingMessage = document.createElement('div')
    loadingMessage.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:rgba(0,0,0,0.7);color:#fff;padding:10px 20px;border-radius:4px;z-index:9999;'
    loadingMessage.textContent = '文档加载中...'
    document.body.appendChild(loadingMessage)

    // 创建FormData对象发送文件到后端
    const formData = new FormData()
    formData.append('file', file)

    // 发送到后端API
    fetch(`${API_BASE_URL}/documents/upload/`, {
      method: 'POST',
      body: formData
    })
    .then(response => {
      if (!response.ok) {
        throw new Error(`上传失败: ${response.status} ${response.statusText}`)
      }
      return response.json()
    })
    .then(data => {
      console.log('上传成功，后端返回:', data)

      // 检查返回数据结构
      if (data && data.analysis && data.analysis.content) {
        try {
          // 将返回的文档内容设置到编辑器
          // 构造编辑器需要的数据格式
          const editorData = {
            header: [],
            main: data.analysis.content,
            footer: [
              {
                value: 'canvas-editor',
                size: 12
              }
            ]
          }

          // 使用executeSetValue方法设置编辑器内容
          if (typeof instance.command.executeSetValue === 'function') {
            instance.command.executeSetValue(editorData)
            console.log('文档内容已加载到编辑器')
          } else {
            throw new Error('编辑器不支持executeSetValue方法，请检查API版本')
          }
        } catch (err) {
          const error = err as Error
          console.error('设置编辑器内容失败:', error)
          throw error
        }
      } else {
        throw new Error('返回数据格式不正确或缺少内容')
      }
    })
    .catch(err => {
      const error = err as Error
      console.error('处理上传文件失败:', error)
      alert('处理上传文件失败: ' + (error.message || '未知错误'))
    })
    .finally(() => {
      // 移除加载提示
      document.body.removeChild(loadingMessage)
      // 清空上传控件的值，以便可以重复上传相同文件
      uploadFileDom.value = ''
    })
  }

  // 下载文档功能
  const downloadDom = document.querySelector<HTMLDivElement>('.menu-item__download')!
  downloadDom.onclick = function () {
    console.log('download')

    // 显示下载选项弹窗
    const optionsDialog = document.createElement('div')
    optionsDialog.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;padding:20px;border-radius:8px;z-index:9999;box-shadow:0 2px 10px rgba(0,0,0,0.2);min-width:250px;'

    optionsDialog.innerHTML = `
      <h3 style="margin-top:0;margin-bottom:15px;font-size:16px;">选择导出格式</h3>
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button id="download-docx" style="padding:8px 12px;cursor:pointer;background:#4285F4;color:white;border:none;border-radius:4px;">导出为DOCX</button>
        <button id="download-json" style="padding:8px 12px;cursor:pointer;background:#f1f1f1;border:1px solid #ccc;border-radius:4px;">导出为JSON</button>
      </div>
      <div style="text-align:right;margin-top:15px;">
        <button id="download-cancel" style="padding:5px 10px;cursor:pointer;background:transparent;border:1px solid #ccc;border-radius:4px;">取消</button>
      </div>
    `

    document.body.appendChild(optionsDialog)

    // 关闭弹窗
    const closeDialog = () => {
      document.body.removeChild(optionsDialog)
    }

    // 设置按钮事件
    document.getElementById('download-cancel')!.onclick = closeDialog

    // 导出为DOCX
    document.getElementById('download-docx')!.onclick = () => {
      closeDialog()

      // 显示加载提示
      const loadingMessage = document.createElement('div')
      loadingMessage.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:rgba(0,0,0,0.7);color:#fff;padding:10px 20px;border-radius:4px;z-index:9999;'
      loadingMessage.textContent = '正在生成DOCX文档...'
      document.body.appendChild(loadingMessage)

      try {
        // 获取编辑器内容
        const content = instance.command.getValue()

        // 构建请求数据
        const requestData = {
          content: content.data,
          options: {
            format: 'docx',
            fileName: 'document-export.docx'
          }
        }

        // 发送到后端API
        fetch(`${API_BASE_URL}/documents/export/`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(requestData)
        })
        .then(response => {
          if (!response.ok) {
            throw new Error(`导出失败: ${response.status} ${response.statusText}`)
          }
          return response.blob()
        })
        .then(blob => {
          // 创建下载链接
          const url = URL.createObjectURL(blob)
          const a = document.createElement('a')
          a.href = url
          a.download = 'document-export.docx'
          a.click()
          URL.revokeObjectURL(url)
          console.log('DOCX文档已下载')
        })
        .catch(err => {
          const error = err as Error
          console.error('导出DOCX失败:', error)
          alert('导出DOCX失败: ' + (error.message || '未知错误'))

          // 降级到JSON导出
          console.log('尝试降级到JSON导出...')
          exportAsJSON()
        })
        .finally(() => {
          document.body.removeChild(loadingMessage)
        })
      } catch (err) {
        const error = err as Error
        console.error('准备DOCX导出失败:', error)
        document.body.removeChild(loadingMessage)
        alert('准备DOCX导出失败: ' + (error.message || '未知错误'))
      }
    }

    // 导出为JSON
    const exportAsJSON = () => {
      try {
        const data = instance.command.getValue()
        const jsonString = JSON.stringify(data, null, 2)
        const blob = new Blob([jsonString], { type: 'application/json' })
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = 'document-content.json'
        a.click()
        URL.revokeObjectURL(url)
        console.log('JSON内容已下载')
        return true
      } catch (err) {
        const error = err as Error
        console.error('导出JSON失败:', error)
        alert('导出JSON失败: ' + (error.message || '未知错误'))
        return false
      }
    }

    document.getElementById('download-json')!.onclick = () => {
      closeDialog()
      exportAsJSON()
    }
  }

  // 6. 目录显隐 | 页面模式 | 纸张缩放 | 纸张大小 | 纸张方向 | 页边距 | 全屏 | 设置
  const editorOptionDom =
    document.querySelector<HTMLDivElement>('.editor-option')!
  editorOptionDom.onclick = function () {
    const options = instance.command.getOptions()
    new Dialog({
      title: '编辑器配置',
      data: [
        {
          type: 'textarea',
          name: 'option',
          width: 350,
          height: 300,
          required: true,
          value: JSON.stringify(options, null, 2),
          placeholder: '请输入编辑器配置'
        }
      ],
      onConfirm: payload => {
        const newOptionValue = payload.find(p => p.name === 'option')?.value
        if (!newOptionValue) return
        const newOption = JSON.parse(newOptionValue)
        instance.command.executeUpdateOptions(newOption)
      }
    })
  }

  async function updateCatalog() {
    const catalog = await instance.command.getCatalog()
    const catalogMainDom =
      document.querySelector<HTMLDivElement>('.catalog__main')!
    catalogMainDom.innerHTML = ''
    if (catalog) {
      const appendCatalog = (
        parent: HTMLDivElement,
        catalogItems: ICatalogItem[]
      ) => {
        for (let c = 0; c < catalogItems.length; c++) {
          const catalogItem = catalogItems[c]
          const catalogItemDom = document.createElement('div')
          catalogItemDom.classList.add('catalog-item')
          // 渲染
          const catalogItemContentDom = document.createElement('div')
          catalogItemContentDom.classList.add('catalog-item__content')
          const catalogItemContentSpanDom = document.createElement('span')
          catalogItemContentSpanDom.innerText = catalogItem.name
          catalogItemContentDom.append(catalogItemContentSpanDom)
          // 定位
          catalogItemContentDom.onclick = () => {
            instance.command.executeLocationCatalog(catalogItem.id)
          }
          catalogItemDom.append(catalogItemContentDom)
          if (catalogItem.subCatalog && catalogItem.subCatalog.length) {
            appendCatalog(catalogItemDom, catalogItem.subCatalog)
          }
          // 追加
          parent.append(catalogItemDom)
        }
      }
      appendCatalog(catalogMainDom, catalog)
    }
  }
  let isCatalogShow = true
  const catalogDom = document.querySelector<HTMLElement>('.catalog')!
  const catalogModeDom =
    document.querySelector<HTMLDivElement>('.catalog-mode')!
  const catalogHeaderCloseDom = document.querySelector<HTMLDivElement>(
    '.catalog__header__close'
  )!
  const switchCatalog = () => {
    isCatalogShow = !isCatalogShow
    if (isCatalogShow) {
      catalogDom.style.display = 'block'
      updateCatalog()
    } else {
      catalogDom.style.display = 'none'
    }
  }
  catalogModeDom.onclick = switchCatalog
  catalogHeaderCloseDom.onclick = switchCatalog

  const pageModeDom = document.querySelector<HTMLDivElement>('.page-mode')!
  const pageModeOptionsDom =
    pageModeDom.querySelector<HTMLDivElement>('.options')!
  pageModeDom.onclick = function () {
    pageModeOptionsDom.classList.toggle('visible')
  }
  pageModeOptionsDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    instance.command.executePageMode(<PageMode>li.dataset.pageMode!)
  }

  document.querySelector<HTMLDivElement>('.page-scale-percentage')!.onclick =
    function () {
      console.log('page-scale-recovery')
      instance.command.executePageScaleRecovery()
    }

  document.querySelector<HTMLDivElement>('.page-scale-minus')!.onclick =
    function () {
      console.log('page-scale-minus')
      instance.command.executePageScaleMinus()
    }

  document.querySelector<HTMLDivElement>('.page-scale-add')!.onclick =
    function () {
      console.log('page-scale-add')
      instance.command.executePageScaleAdd()
    }

  // 纸张大小
  const paperSizeDom = document.querySelector<HTMLDivElement>('.paper-size')!
  const paperSizeDomOptionsDom =
    paperSizeDom.querySelector<HTMLDivElement>('.options')!
  paperSizeDom.onclick = function () {
    paperSizeDomOptionsDom.classList.toggle('visible')
  }
  paperSizeDomOptionsDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    const paperType = li.dataset.paperSize!
    const [width, height] = paperType.split('*').map(Number)
    instance.command.executePaperSize(width, height)
    // 纸张状态回显
    paperSizeDomOptionsDom
      .querySelectorAll('li')
      .forEach(child => child.classList.remove('active'))
    li.classList.add('active')
  }

  // 纸张方向
  const paperDirectionDom =
    document.querySelector<HTMLDivElement>('.paper-direction')!
  const paperDirectionDomOptionsDom =
    paperDirectionDom.querySelector<HTMLDivElement>('.options')!
  paperDirectionDom.onclick = function () {
    paperDirectionDomOptionsDom.classList.toggle('visible')
  }
  paperDirectionDomOptionsDom.onclick = function (evt) {
    const li = evt.target as HTMLLIElement
    const paperDirection = li.dataset.paperDirection!
    instance.command.executePaperDirection(<PaperDirection>paperDirection)
    // 纸张方向状态回显
    paperDirectionDomOptionsDom
      .querySelectorAll('li')
      .forEach(child => child.classList.remove('active'))
    li.classList.add('active')
  }

  // 页面边距
  const paperMarginDom =
    document.querySelector<HTMLDivElement>('.paper-margin')!
  paperMarginDom.onclick = function () {
    const [topMargin, rightMargin, bottomMargin, leftMargin] =
      instance.command.getPaperMargin()
    new Dialog({
      title: '页边距',
      data: [
        {
          type: 'text',
          label: '上边距',
          name: 'top',
          required: true,
          value: `${topMargin}`,
          placeholder: '请输入上边距'
        },
        {
          type: 'text',
          label: '下边距',
          name: 'bottom',
          required: true,
          value: `${bottomMargin}`,
          placeholder: '请输入下边距'
        },
        {
          type: 'text',
          label: '左边距',
          name: 'left',
          required: true,
          value: `${leftMargin}`,
          placeholder: '请输入左边距'
        },
        {
          type: 'text',
          label: '右边距',
          name: 'right',
          required: true,
          value: `${rightMargin}`,
          placeholder: '请输入右边距'
        }
      ],
      onConfirm: payload => {
        const top = payload.find(p => p.name === 'top')?.value
        if (!top) return
        const bottom = payload.find(p => p.name === 'bottom')?.value
        if (!bottom) return
        const left = payload.find(p => p.name === 'left')?.value
        if (!left) return
        const right = payload.find(p => p.name === 'right')?.value
        if (!right) return
        instance.command.executeSetPaperMargin([
          Number(top),
          Number(right),
          Number(bottom),
          Number(left)
        ])
      }
    })
  }

  // 全屏
  const fullscreenDom = document.querySelector<HTMLDivElement>('.fullscreen')!
  fullscreenDom.onclick = toggleFullscreen
  window.addEventListener('keydown', evt => {
    if (evt.key === 'F11') {
      toggleFullscreen()
      evt.preventDefault()
    }
  })
  document.addEventListener('fullscreenchange', () => {
    fullscreenDom.classList.toggle('exist')
  })
  function toggleFullscreen() {
    console.log('fullscreen')
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen()
    } else {
      document.exitFullscreen()
    }
  }

  // 7. 编辑器使用模式
  let modeIndex = 0
  const modeList = [
    {
      mode: EditorMode.EDIT,
      name: '编辑模式'
    },
    {
      mode: EditorMode.CLEAN,
      name: '清洁模式'
    },
    {
      mode: EditorMode.READONLY,
      name: '只读模式'
    },
    {
      mode: EditorMode.FORM,
      name: '表单模式'
    },
    {
      mode: EditorMode.PRINT,
      name: '打印模式'
    },
    {
      mode: EditorMode.DESIGN,
      name: '设计模式'
    }
  ]
  const modeElement = document.querySelector<HTMLDivElement>('.editor-mode')!
  modeElement.onclick = function () {
    // 模式选择循环
    modeIndex === modeList.length - 1 ? (modeIndex = 0) : modeIndex++
    // 设置模式
    const { name, mode } = modeList[modeIndex]
    modeElement.innerText = name
    instance.command.executeMode(mode)
    // 设置菜单栏权限视觉反馈
    const isReadonly = mode === EditorMode.READONLY
    const enableMenuList = ['search', 'print']
    document.querySelectorAll<HTMLDivElement>('.menu-item>div').forEach(dom => {
      const menu = dom.dataset.menu
      isReadonly && (!menu || !enableMenuList.includes(menu))
        ? dom.classList.add('disable')
        : dom.classList.remove('disable')
    })
  }

  // 模拟批注
  const commentDom = document.querySelector<HTMLDivElement>('.comment')!
  async function updateComment() {
    const groupIds = await instance.command.getGroupIds()
    for (const comment of commentList) {
      const activeCommentDom = commentDom.querySelector<HTMLDivElement>(
        `.comment-item[data-id='${comment.id}']`
      )
      // 编辑器是否存在对应成组id
      if (groupIds.includes(comment.id)) {
        // 当前dom是否存在-不存在则追加
        if (!activeCommentDom) {
          const commentItem = document.createElement('div')
          commentItem.classList.add('comment-item')
          commentItem.setAttribute('data-id', comment.id)
          commentItem.onclick = () => {
            instance.command.executeLocationGroup(comment.id)
          }
          commentDom.append(commentItem)
          // 选区信息
          const commentItemTitle = document.createElement('div')
          commentItemTitle.classList.add('comment-item__title')
          commentItemTitle.append(document.createElement('span'))
          const commentItemTitleContent = document.createElement('span')
          commentItemTitleContent.innerText = comment.rangeText
          commentItemTitle.append(commentItemTitleContent)
          const closeDom = document.createElement('i')
          closeDom.onclick = () => {
            instance.command.executeDeleteGroup(comment.id)
          }
          commentItemTitle.append(closeDom)
          commentItem.append(commentItemTitle)
          // 基础信息
          const commentItemInfo = document.createElement('div')
          commentItemInfo.classList.add('comment-item__info')
          const commentItemInfoName = document.createElement('span')
          commentItemInfoName.innerText = comment.userName
          const commentItemInfoDate = document.createElement('span')
          commentItemInfoDate.innerText = comment.createdDate
          commentItemInfo.append(commentItemInfoName)
          commentItemInfo.append(commentItemInfoDate)
          commentItem.append(commentItemInfo)
          // 详细评论
          const commentItemContent = document.createElement('div')
          commentItemContent.classList.add('comment-item__content')
          commentItemContent.innerText = comment.content
          commentItem.append(commentItemContent)
          commentDom.append(commentItem)
        }
      } else {
        // 编辑器内不存在对应成组id则dom则移除
        activeCommentDom?.remove()
      }
    }
  }
  // 8. 内部事件监听
  instance.listener.rangeStyleChange = function (payload) {
    // 控件类型
    payload.type === ElementType.SUBSCRIPT
      ? subscriptDom.classList.add('active')
      : subscriptDom.classList.remove('active')
    payload.type === ElementType.SUPERSCRIPT
      ? superscriptDom.classList.add('active')
      : superscriptDom.classList.remove('active')
    payload.type === ElementType.SEPARATOR
      ? separatorDom.classList.add('active')
      : separatorDom.classList.remove('active')
    separatorOptionDom
      .querySelectorAll('li')
      .forEach(li => li.classList.remove('active'))
    if (payload.type === ElementType.SEPARATOR) {
      const separator = payload.dashArray.join(',') || '0,0'
      const curSeparatorDom = separatorOptionDom.querySelector<HTMLLIElement>(
        `[data-separator='${separator}']`
      )!
      if (curSeparatorDom) {
        curSeparatorDom.classList.add('active')
      }
    }

    // 富文本
    fontOptionDom
      .querySelectorAll<HTMLLIElement>('li')
      .forEach(li => li.classList.remove('active'))
    const curFontDom = fontOptionDom.querySelector<HTMLLIElement>(
      `[data-family='${payload.font}']`
    )
    if (curFontDom) {
      fontSelectDom.innerText = curFontDom.innerText
      fontSelectDom.style.fontFamily = payload.font
      curFontDom.classList.add('active')
    }
    sizeOptionDom
      .querySelectorAll<HTMLLIElement>('li')
      .forEach(li => li.classList.remove('active'))
    const curSizeDom = sizeOptionDom.querySelector<HTMLLIElement>(
      `[data-size='${payload.size}']`
    )
    if (curSizeDom) {
      sizeSelectDom.innerText = curSizeDom.innerText
      curSizeDom.classList.add('active')
    } else {
      sizeSelectDom.innerText = `${payload.size}`
    }
    payload.bold
      ? boldDom.classList.add('active')
      : boldDom.classList.remove('active')
    payload.italic
      ? italicDom.classList.add('active')
      : italicDom.classList.remove('active')
    payload.underline
      ? underlineDom.classList.add('active')
      : underlineDom.classList.remove('active')
    payload.strikeout
      ? strikeoutDom.classList.add('active')
      : strikeoutDom.classList.remove('active')
    if (payload.color) {
      colorDom.classList.add('active')
      colorControlDom.value = payload.color
      colorSpanDom.style.backgroundColor = payload.color
    } else {
      colorDom.classList.remove('active')
      colorControlDom.value = '#000000'
      colorSpanDom.style.backgroundColor = '#000000'
    }
    if (payload.highlight) {
      highlightDom.classList.add('active')
      highlightControlDom.value = payload.highlight
      highlightSpanDom.style.backgroundColor = payload.highlight
    } else {
      highlightDom.classList.remove('active')
      highlightControlDom.value = '#ffff00'
      highlightSpanDom.style.backgroundColor = '#ffff00'
    }

    // 行布局
    leftDom.classList.remove('active')
    centerDom.classList.remove('active')
    rightDom.classList.remove('active')
    alignmentDom.classList.remove('active')
    justifyDom.classList.remove('active')
    if (payload.rowFlex && payload.rowFlex === 'right') {
      rightDom.classList.add('active')
    } else if (payload.rowFlex && payload.rowFlex === 'center') {
      centerDom.classList.add('active')
    } else if (payload.rowFlex && payload.rowFlex === 'alignment') {
      alignmentDom.classList.add('active')
    } else if (payload.rowFlex && payload.rowFlex === 'justify') {
      justifyDom.classList.add('active')
    } else {
      leftDom.classList.add('active')
    }

    // 行间距
    rowOptionDom
      .querySelectorAll<HTMLLIElement>('li')
      .forEach(li => li.classList.remove('active'))
    const curRowMarginDom = rowOptionDom.querySelector<HTMLLIElement>(
      `[data-rowmargin='${payload.rowMargin}']`
    )!
    curRowMarginDom.classList.add('active')

    // 功能
    payload.undo
      ? undoDom.classList.remove('no-allow')
      : undoDom.classList.add('no-allow')
    payload.redo
      ? redoDom.classList.remove('no-allow')
      : redoDom.classList.add('no-allow')
    payload.painter
      ? painterDom.classList.add('active')
      : painterDom.classList.remove('active')

    // 标题
    titleOptionDom
      .querySelectorAll<HTMLLIElement>('li')
      .forEach(li => li.classList.remove('active'))
    if (payload.level) {
      const curTitleDom = titleOptionDom.querySelector<HTMLLIElement>(
        `[data-level='${payload.level}']`
      )!
      titleSelectDom.innerText = curTitleDom.innerText
      curTitleDom.classList.add('active')
    } else {
      titleSelectDom.innerText = '正文'
      titleOptionDom.querySelector('li:first-child')!.classList.add('active')
    }

    // 列表
    listOptionDom
      .querySelectorAll<HTMLLIElement>('li')
      .forEach(li => li.classList.remove('active'))
    if (payload.listType) {
      listDom.classList.add('active')
      const listType = payload.listType
      const listStyle =
        payload.listType === ListType.OL ? ListStyle.DECIMAL : payload.listType
      const curListDom = listOptionDom.querySelector<HTMLLIElement>(
        `[data-list-type='${listType}'][data-list-style='${listStyle}']`
      )
      if (curListDom) {
        curListDom.classList.add('active')
      }
    } else {
      listDom.classList.remove('active')
    }

    // 批注
    commentDom
      .querySelectorAll<HTMLDivElement>('.comment-item')
      .forEach(commentItemDom => {
        commentItemDom.classList.remove('active')
      })
    if (payload.groupIds) {
      const [id] = payload.groupIds
      const activeCommentDom = commentDom.querySelector<HTMLDivElement>(
        `.comment-item[data-id='${id}']`
      )
      if (activeCommentDom) {
        activeCommentDom.classList.add('active')
        scrollIntoView(commentDom, activeCommentDom)
      }
    }
  }

  instance.listener.visiblePageNoListChange = function (payload) {
    const text = payload.map(i => i + 1).join('、')
    document.querySelector<HTMLSpanElement>('.page-no-list')!.innerText = text
  }

  instance.listener.pageSizeChange = function (payload) {
    document.querySelector<HTMLSpanElement>(
      '.page-size'
    )!.innerText = `${payload}`
  }

  instance.listener.intersectionPageNoChange = function (payload) {
    document.querySelector<HTMLSpanElement>('.page-no')!.innerText = `${
      payload + 1
    }`
  }

  instance.listener.pageScaleChange = function (payload) {
    document.querySelector<HTMLSpanElement>(
      '.page-scale-percentage'
    )!.innerText = `${Math.floor(payload * 10 * 10)}%`
  }

  instance.listener.controlChange = function (payload) {
    const disableMenusInControlContext = [
      'table',
      'hyperlink',
      'separator',
      'page-break',
      'control'
    ]
    // 菜单操作权限
    disableMenusInControlContext.forEach(menu => {
      const menuDom = document.querySelector<HTMLDivElement>(
        `.menu-item__${menu}`
      )!
      payload.state === ControlState.ACTIVE
        ? menuDom.classList.add('disable')
        : menuDom.classList.remove('disable')
    })
  }

  instance.listener.pageModeChange = function (payload) {
    const activeMode = pageModeOptionsDom.querySelector<HTMLLIElement>(
      `[data-page-mode='${payload}']`
    )!
    pageModeOptionsDom
      .querySelectorAll('li')
      .forEach(li => li.classList.remove('active'))
    activeMode.classList.add('active')
  }

  const handleContentChange = async function () {
    // 字数
    const wordCount = await instance.command.getWordCount()
    document.querySelector<HTMLSpanElement>('.word-count')!.innerText = `${
      wordCount || 0
    }`
    // 目录
    if (isCatalogShow) {
      nextTick(() => {
        updateCatalog()
      })
    }
    // 批注
    nextTick(() => {
      updateComment()
    })
  }
  instance.listener.contentChange = debounce(handleContentChange, 200)
  handleContentChange()

  instance.listener.saved = function (payload) {
    console.log('elementList: ', payload)
  }

  // 9. 右键菜单注册
  instance.register.contextMenuList([
    {
      name: '批注',
      when: payload => {
        return (
          !payload.isReadonly &&
          payload.editorHasSelection &&
          payload.zone === EditorZone.MAIN
        )
      },
      callback: (command: Command) => {
        new Dialog({
          title: '批注',
          data: [
            {
              type: 'textarea',
              label: '批注',
              height: 100,
              name: 'value',
              required: true,
              placeholder: '请输入批注'
            }
          ],
          onConfirm: payload => {
            const value = payload.find(p => p.name === 'value')?.value
            if (!value) return
            const groupId = command.executeSetGroup()
            if (!groupId) return
            commentList.push({
              id: groupId,
              content: value,
              userName: 'Hufe',
              rangeText: command.getRangeText(),
              createdDate: new Date().toLocaleString()
            })
          }
        })
      }
    },
    {
      name: '签名',
      icon: 'signature',
      when: payload => {
        return !payload.isReadonly && payload.editorTextFocus
      },
      callback: (command: Command) => {
        new Signature({
          onConfirm(payload) {
            if (!payload) return
            const { value, width, height } = payload
            if (!value || !width || !height) return
            command.executeInsertElementList([
              {
                value,
                width,
                height,
                type: ElementType.IMAGE
              }
            ])
          }
        })
      }
    },
    {
      name: '格式整理',
      icon: 'word-tool',
      when: payload => {
        return !payload.isReadonly
      },
      callback: (command: Command) => {
        command.executeWordTool()
      }
    }
  ])

  // 10. 快捷键注册
  instance.register.shortcutList([
    {
      key: KeyMap.P,
      mod: true,
      isGlobal: true,
      callback: (command: Command) => {
        command.executePrint()
      }
    },
    {
      key: KeyMap.F,
      mod: true,
      isGlobal: true,
      callback: (command: Command) => {
        const text = command.getRangeText()
        searchDom.click()
        if (text) {
          searchInputDom.value = text
          instance.command.executeSearch(text)
          setSearchResult()
        }
      }
    },
    {
      key: KeyMap.MINUS,
      ctrl: true,
      isGlobal: true,
      callback: (command: Command) => {
        command.executePageScaleMinus()
      }
    },
    {
      key: KeyMap.EQUAL,
      ctrl: true,
      isGlobal: true,
      callback: (command: Command) => {
        command.executePageScaleAdd()
      }
    },
    {
      key: KeyMap.ZERO,
      ctrl: true,
      isGlobal: true,
      callback: (command: Command) => {
        command.executePageScaleRecovery()
      }
    }
  ])
}
