using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace REghZY.SelectionSurrounder {
    [Export(typeof(IVsTextViewCreationListener))]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    [ContentType("text")]
    internal sealed class VsTextViewListener : IVsTextViewCreationListener {
        [Import]
        internal IVsEditorAdaptersFactoryService AdapterService = null;

        public void VsTextViewCreated(IVsTextView textViewAdapter) {
            ITextView textView = AdapterService.GetWpfTextView(textViewAdapter);
            if (textView == null)
                return;

            textView.Properties.GetOrCreateSingletonProperty(() => new TypeCharFilter(textViewAdapter, textView));
        }
    }

    internal sealed class TypeCharFilter : IOleCommandTarget {
        private readonly IOleCommandTarget nextCommandHandler;
        private readonly ITextView textView;
        internal int TypedChars { get; set; }

        /// <summary>
        /// Add this filter to the chain of Command Filters
        /// </summary>
        internal TypeCharFilter(IVsTextView adapter, ITextView textView) {
            this.textView = textView;
            adapter.AddCommandFilter(this, out nextCommandHandler);
        }

        /// <summary>
        /// Get user input
        /// </summary>
        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut) {
            int hr;
            if (TryGetTypedChar(pguidCmdGroup, nCmdID, pvaIn, out char typedChar)) {
                ITextSelection selection = textView.Selection;
                if (HandleKey(selection, typedChar)) {
                    return VSConstants.S_OK;
                }
                else {
                    hr = nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                    return hr;
                }
            }
            else {
                hr = nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                return hr;
            }
        }

        /// <summary>
        /// Handles a key press
        /// </summary>
        /// <returns>True if this extension has handled the key press, aka modified the text buffer in some way. False if not</returns>
        public bool HandleKey(ITextSelection selection, char input) {
            if (selection == null) {
                return false;
            }
            if (selection.IsEmpty) {
                return false;
            }

            if (selection.Mode == TextSelectionMode.Stream) {
                int selectionLength = selection.End.Position.Position - selection.Start.Position.Position;
                if (Keyboard.IsKeyDown(Key.LeftShift)) {
                    if (input == '(') {
                        InsertBetween(selection, "(", ")", selectionLength);
                    }
                    else if (input == '{') {
                        InsertBetween(selection, "{", "}", selectionLength);
                    }
                    else if (input == '<') {
                        InsertBetween(selection, "<", ">", selectionLength);
                    }
                    else if (input == '\"') {
                        InsertBetween(selection, "\"", "\"", selectionLength);
                    }
                    else {
                        return false;
                    }

                    return true;
                }
                else {
                    if (input == '[') {
                        InsertBetween(selection, "[", "]", selectionLength);
                    }
                    else if (input == '\'') {
                        InsertBetween(selection, "\'", "\'", selectionLength);
                    }
                    else {
                        return false;
                    }

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Inserts the given a and b string before and after the given selection, and then re-selects that text using the given selection length
        /// </summary>
        public void InsertBetween(ITextSelection selection, string a, string b, int selectionLength) {
            textView.TextBuffer.Insert(selection.Start.Position.Position, a);
            textView.TextBuffer.Insert(selection.End.Position.Position, b);
            selection.Select(new SnapshotSpan(selection.Start.Position, selectionLength), false);
        }

        /// <summary>
        /// Public access to IOleCommandTarget.QueryStatus() function
        /// </summary>
        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText) {
            return nextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        /// <summary>
        /// Try to get the keypress value. Returns 0 if attempt fails
        /// </summary>
        /// <param name="typedChar">Outputs the value of the typed char</param>
        /// <returns>Boolean reporting success or failure of operation</returns>
        bool TryGetTypedChar(Guid cmdGroup, uint nCmdID, IntPtr pvaIn, out char typedChar) {
            if (cmdGroup != VSConstants.VSStd2K || nCmdID != (uint)VSConstants.VSStd2KCmdID.TYPECHAR) {
                typedChar = char.MinValue;
                return false;
            }

            typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
            return true;
        }
    }
}
