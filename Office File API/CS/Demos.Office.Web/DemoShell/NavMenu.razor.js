export async function ensureNodeVisible(selectedNodeId, treeViewId) {
    var selectedNode = document.getElementById(selectedNodeId);
    if (selectedNode) {
        var selectedNodeRect = selectedNode.getBoundingClientRect();
        var scrollViewerContent = document.getElementById(treeViewId)?.parentElement;
        if (scrollViewerContent) {
            var scrollViewerContentRect = scrollViewerContent.getBoundingClientRect();
            if (selectedNodeRect.top < scrollViewerContentRect.top || selectedNodeRect.bottom > scrollViewerContentRect.bottom)
                selectedNode.scrollIntoView();
        }
    }
}
