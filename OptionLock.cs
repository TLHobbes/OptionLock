using System;
using System.Collections.Generic;
using System.Windows.Forms;

using KeePass.Plugins;

namespace OptionLock
{
  public sealed class OptionLockExt : Plugin
  {
    IPluginHost m_Host;

    // KeePass' FileLock menu item
    ToolStripItem m_FileLockItem;

    // UI made to be accessible only when there is an unlocked database.
    LinkedList<ToolStripItem> m_UnlockedDbItems;

    // UI made to be accessible only when there is an unlocked database
    // or when there are no opened documents.
    LinkedList<ToolStripItem> m_NoDocItems;

    #region Setup and Teardown

    public override bool Initialize(IPluginHost host)
    {
      m_Host = host;

      InitializeItems();

      // Detect lock state changes by the opening/closing of files
      m_Host.MainWindow.FileOpened += MainWindow_FileOpened;
      m_Host.MainWindow.FileClosed += MainWindow_FileClosed;

      return base.Initialize(host);
    }

    public override void Terminate()
    {
      CleanupItems();
      base.Terminate();
    }

    // Find, initialize and track KeePass' UI items of interest
    void InitializeItems()
    {
      m_UnlockedDbItems = new LinkedList<ToolStripItem>();
      m_NoDocItems = new LinkedList<ToolStripItem>();

      // TrayContextMenu
      foreach (var item in m_Host.MainWindow.TrayContextMenu.Items)
      {
        var stripItem = item as ToolStripItem;
        if (stripItem != null)
        {
          switch (stripItem.Name)
          {
            /// --- BEGIN ALLOW LIST --- ///
            case "m_ctxTrayTray":
            case "m_ctxTrayLock":
            case "m_ctxTrayFileExit":
                // Do not add these allowed items
                break;
            /// --- END ALLOW LIST --- ///
                
            default:
                // Add everything else, which will include any future added menu items from newer KeePass 2 releases
                m_UnlockedDbItems.AddFirst(stripItem);
                break;
          }
        }
      }

      // MainMenu
      foreach (var mainItem in m_Host.MainWindow.MainMenu.Items)
      {
        var dropDown = mainItem as ToolStripDropDownItem;
        if (dropDown != null)
        {
          foreach (var item in dropDown.DropDownItems)
          {
            var stripItem = item as ToolStripItem;
            if (stripItem != null)
            {
              switch (stripItem.Name)
              {
                // File
                case "m_menuFileOpen":
                case "m_menuFileRecent":
                  m_NoDocItems.AddFirst(stripItem);
                  break;

                /// --- BEGIN ALLOW LIST --- ///
                // File
                case "m_menuFileLock":
                  m_FileLockItem = stripItem;
                  m_FileLockItem.EnabledChanged += FileLockMenuItem_EnabledChanged;
                  break;
                case "m_menuFileExit":
                // Help
                case "m_menuHelpContents":
                case "m_menuHelpWebsite":
                case "m_menuHelpDonate":
                case "m_menuHelpAbout":
                  // Do not add these allowed items
                  break;
                /// --- END ALLOW LIST --- ///
                
                default:
                  // Add everything else, which will include any future added menu items from newer KeePass 2 releases
                  m_UnlockedDbItems.AddFirst(stripItem);
                  break;
              }
            }
          }
        }
      }

      // CustomToolStrip
      int toolIndex = m_Host.MainWindow.Controls.IndexOfKey("m_toolMain");
      var toolStrip = m_Host.MainWindow.Controls[toolIndex] as KeePass.UI.CustomToolStripEx;
      foreach (var item in toolStrip.Items)
      {
        var stripItem = item as ToolStripItem;
        if (stripItem != null)
        {
          switch (stripItem.Name)
          {
            /// --- BEGIN ALLOW LIST --- ///
            case "m_tbOpenDatabase":
              m_NoDocItems.AddFirst(stripItem);
              break;
            case "m_tbLockWorkspace":
              // Do not add these allowed items
              break;
            /// --- END ALLOW LIST --- ///
                        
            default:
              m_UnlockedDbItems.AddFirst(stripItem);
              break;
          }
        }
      }

      // Initialize enabled states of items and track changes
      bool isUnlocked = IsAtLeastOneFileOpenAndUnlocked();
      foreach (var item in m_UnlockedDbItems)
      {
        item.Enabled = isUnlocked;
        item.EnabledChanged += UnlockedDbMenuItem_EnabledChanged;
      }
      foreach (var item in m_NoDocItems)
      {
        item.Enabled = HasNoDocs;
        item.EnabledChanged += NoDbMenuItem_EnabledChanged;
      }
    }

    // Dereference KeePass's UI items
    void CleanupItems()
    {
      m_FileLockItem.EnabledChanged -= FileLockMenuItem_EnabledChanged;
      m_FileLockItem = null;

      foreach (var item in m_UnlockedDbItems)
      {
        item.EnabledChanged -= UnlockedDbMenuItem_EnabledChanged;
        item.Enabled = true;
      }
      m_UnlockedDbItems.Clear();

      foreach (var item in m_NoDocItems)
      {
        item.EnabledChanged -= NoDbMenuItem_EnabledChanged;
        item.Enabled = true;
      }
      m_NoDocItems.Clear();
    }

    #endregion // Setup and Teardown

    #region Helpers

    bool HasDocs { get { return m_FileLockItem.Enabled; } }
    bool HasNoDocs { get { return !HasDocs; } }

    bool IsAtLeastOneFileOpenAndUnlocked()
    {
      var mainForm = m_Host.MainWindow;
      foreach (var doc in mainForm.DocumentManager.Documents)
      {
        if (doc.Database.IsOpen && !mainForm.IsFileLocked(doc))
        {
          return true;
        }
      }
      return false;
    }

    void EnableItem(bool enabled, ToolStripItem item, EventHandler enabledChangedHandler)
    {
      if (item.Enabled != enabled)
      {
        item.EnabledChanged -= enabledChangedHandler;
        item.Enabled = enabled;
        item.EnabledChanged += enabledChangedHandler;
      }
    }

    void EnableNoDbItem(bool enabled, ToolStripItem item)
    {
      EnableItem(enabled, item, NoDbMenuItem_EnabledChanged);
    }

    void EnableUnlockedDbItem(bool enabled, ToolStripItem item)
    {
      EnableItem(enabled, item, UnlockedDbMenuItem_EnabledChanged);
    }

    #endregion // Helpers

    #region Event Handlers

    // KeePass opened a file/database
    void MainWindow_FileOpened(object sender, KeePass.Forms.FileOpenedEventArgs e)
    {
      if (IsAtLeastOneFileOpenAndUnlocked())
      {
        // An unlocked database exists so enable all tracked items.
        foreach (var item in m_UnlockedDbItems)
        {
          EnableUnlockedDbItem(true, item);
        }
        foreach (var item in m_NoDocItems)
        {
          EnableNoDbItem(true, item);
        }
      }
    }

    // KeePass closed a file/database
    void MainWindow_FileClosed(object sender, KeePass.Forms.FileClosedEventArgs e)
    {
      // Disable tracked items if there is no unlocked database
      if (!IsAtLeastOneFileOpenAndUnlocked())
      {
        // No unlocked databases exist so disable all tracked items.
        foreach (var item in m_UnlockedDbItems)
        {
          EnableUnlockedDbItem(false, item);
        }

        // Except, only disable NoDoc items when there are docs
        if (HasDocs)
        {
          foreach (var item in m_NoDocItems)
          {
            EnableNoDbItem(false, item);
          }
        }        
      }
    }

    // KeePass changed enabled state of its FileLock menu item
    void FileLockMenuItem_EnabledChanged(object sender, EventArgs e)
    {
      // When no docs are open, FileLock menu item is disabled;
      // otherwise, it is enabled. KeePass enables it *before*
      // firing FileOpened and disables it *after* firing FileClosed.
      // Respectively only for first/last of File open/close given
      // that KeePass features multiple concurrent opened files.
      foreach (var item in m_NoDocItems)
      {
        EnableNoDbItem(HasNoDocs, item);
      }
    }

    // Something external changed the state of a tracked NoDoc item
    // so fix its enabled state if needed.
    void NoDbMenuItem_EnabledChanged(object sender, EventArgs e)
    {
      var item = sender as ToolStripItem;
      if (item != null)
      {
        if (HasNoDocs)
        {
          EnableNoDbItem(true, item);
        }
        else
        {
          EnableNoDbItem(IsAtLeastOneFileOpenAndUnlocked(), item);
        }
      }
    }

    // Something external changed the state of a tracked UnlockedDb item
    // so fix its enabled state if needed.
    // KeePass does cause this case to happen as it likes to refresh the
    // enabled state of the "Close" and "Options" UI here and there.
    void UnlockedDbMenuItem_EnabledChanged(object sender, EventArgs e)
    {
      var item = sender as ToolStripItem;
      if (item != null)
      {
        EnableUnlockedDbItem(IsAtLeastOneFileOpenAndUnlocked(), item);
      }
    }

    #endregion // Event Handlers
  }
}
