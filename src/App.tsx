/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Package, 
  Plus, 
  Minus, 
  History, 
  LayoutDashboard, 
  Settings, 
  LogOut,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  TrendingDown,
  Boxes,
  ClipboardList,
  Trash2,
  UserCircle,
  Lock,
  User,
  Warehouse,
  Mail,
  Building2,
  Users,
  Shield,
  CheckCircle,
  ShoppingCart,
  BadgeDollarSign,
  Download,
  AlertTriangle,
  Eye,
  EyeOff,
  Pencil,
  X,
  Search,
  Activity,
  Menu,
  Bell,
  Check,
  FileSpreadsheet,
  Factory,
  PackageCheck,
  Layers,
  Sparkles,
  TrendingUp
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format } from 'date-fns';
import { es } from 'date-fns/locale';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import * as XLSX from 'xlsx';
import { 
  AreaChart, 
  Area, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip as RechartsTooltip, 
  ResponsiveContainer 
} from 'recharts';
import { Product, Batch, Transaction, UserRole, KardexEntry, AuditEntry, TeamMember, AreaPermission, Account, AccountType, AppNotification, RecipeItem, PackagingConfig, ProductionOrderRecord } from './types';
import { processExit, generateUnifiedKardex, simulateExitImpact, TransactionSimulationResult } from './logic/peps';
import { exportKardexWorkbook } from './utils/kardexExcelExport';
import { TransactionConfirmationModal } from './components/TransactionConfirmationModal';
import { authService } from './services/auth';
import { dbService } from './services/db';
import firebaseConfig from '../firebase-applet-config.json';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Initial Mock Data
const INITIAL_PRODUCTS: Product[] = [];
const INITIAL_TEAM: TeamMember[] = [
  { id: '1', name: 'Admin Root', email: 'admin@sied.com', role: 'ADMIN', areas: ['INVENTARIO', 'REPORTES', 'CONFIGURACION'], status: 'active' }
];

export default function App() {
  const [currentUser, setCurrentUser] = useState<Account | null>(() => {
    const saved = localStorage.getItem('sied_current_user');
    return saved ? JSON.parse(saved) : null;
  });

  const [products, setProducts] = useState<Product[]>([]);
  const [team, setTeam] = useState<TeamMember[]>([]);
  const [batches, setBatches] = useState<Batch[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [productionOrders, setProductionOrders] = useState<ProductionOrderRecord[]>([]);
  const [auditLogs, setAuditLogs] = useState<AuditEntry[]>([]);
  const [isSeedingData, setIsSeedingData] = useState(false);
  const [availableCompanies, setAvailableCompanies] = useState<Array<{ id: string; name: string; productCount: number; batchCount: number }>>([]);

  // Fallback to currentUser.id when company account or when linkedCompanyId is missing
  const effectiveCompanyId = currentUser?.linkedCompanyId || (currentUser?.type === 'COMPANY' ? currentUser?.id : (currentUser?.id || ''));

  const [activeTab, setActiveTab] = useState<'dashboard' | 'inventory' | 'purchases' | 'sales' | 'kardex' | 'people' | 'configuration' | 'audit'>('dashboard');
  const [criticalThreshold, setCriticalThreshold] = useState(() => {
    const saved = localStorage.getItem('sied_critical_threshold');
    return saved ? parseFloat(saved) : 0.25;
  });
  const [showRegisterOptions, setShowRegisterOptions] = useState(false);
  const [registerForm, setRegisterForm] = useState<'empresa' | 'usuario' | null>(null);
  const [regFormData, setRegFormData] = useState({ company: '', name: '', email: '', password: '' });
  const [loginFormData, setLoginFormData] = useState({ email: '', password: '' });
  const [isForgotPassword, setIsForgotPassword] = useState(false);
  const [resetEmail, setResetEmail] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showNewPassword, setShowNewPassword] = useState(false);
  const [showConfirmPassword, setShowConfirmPassword] = useState(false);
  const [isResettingPassword, setIsResettingPassword] = useState(false);
  const [resetError, setResetError] = useState('');
  const [resetSuccess, setResetSuccess] = useState(false);
  const [resetOobCode, setResetOobCode] = useState<string | null>(null);

  useEffect(() => {
    // Detect password reset link from email action code
    const params = new URLSearchParams(window.location.search);
    const mode = params.get('mode');
    const oobCode = params.get('oobCode');
    if (mode === 'resetPassword' && oobCode) {
      setResetOobCode(oobCode);
      setIsForgotPassword(true);
    }
  }, []);
  const [showPeopleModal, setShowPeopleModal] = useState(false);
  const [newPerson, setNewPerson] = useState<Partial<TeamMember>>({ areas: [], role: 'OPERATOR' });
  const [selectedProductId, setSelectedProductId] = useState<string | null>(null);
  const [showModal, setShowModal] = useState<'entry' | 'exit' | 'product' | 'purchase' | 'sale' | 'edit_batch' | 'product_history' | 'produce' | null>(null);
  const [selectedBatchId, setSelectedBatchId] = useState<string | null>(null);
  const [productToDelete, setProductToDelete] = useState<Product | null>(null);
  const [productToEdit, setProductToEdit] = useState<Product | null>(null);
  const [editProductForm, setEditProductForm] = useState<{
    name: string;
    description: string;
    category: string;
    type: 'materia prima' | 'producción' | 'producto terminado';
    unit: 'unidad' | 'kilo' | 'litro';
    minStock: number;
    recipe: RecipeItem[];
    packaging: PackagingConfig;
  }>({
    name: '',
    description: '',
    category: '',
    type: 'materia prima',
    unit: 'litro',
    minStock: 0,
    recipe: [],
    packaging: { packagingType: 'ninguno', unitsPerBox: 0, unitsPerBundle: 0 }
  });
  // State for transforming / producing a finished product
  const [produceForm, setProduceForm] = useState<{
    productId: string;
    quantityToProduce: number;
    producedIn: 'unidades' | 'cajas' | 'bultos';
    boxCount: number;
    bundleCount: number;
    notes: string;
    productionDate: string;
  }>({
    productId: '',
    quantityToProduce: 1,
    producedIn: 'unidades',
    boxCount: 0,
    bundleCount: 0,
    notes: '',
    productionDate: new Date().toISOString().split('T')[0]
  });
  const [productRecipe, setProductRecipe] = useState<RecipeItem[]>([]);
  const [productPackaging, setProductPackaging] = useState<PackagingConfig>({
    packagingType: 'ninguno',
    unitsPerBox: 0,
    unitsPerBundle: 0
  });
  const [showLowStockOnly, setShowLowStockOnly] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [kardexPresentationMode, setKardexPresentationMode] = useState(false);
  const [batchToEdit, setBatchToEdit] = useState<Batch | null>(null);
  const [editBatchData, setEditBatchData] = useState({ quantity: 0, cost: 0, reference: '' });
  
  // Edit purchase state
  const [purchaseToEdit, setPurchaseToEdit] = useState<Batch | null>(null);
  const [editPurchaseForm, setEditPurchaseForm] = useState({
    quantity: 0,
    unitCost: 0,
    supplier: '',
    entryDate: '',
    description: ''
  });

  // Edit sale state
  const [saleToEdit, setSaleToEdit] = useState<Transaction | null>(null);
  const [editSaleForm, setEditSaleForm] = useState({
    quantity: 0,
    salePrice: 0,
    date: '',
    reference: '',
    description: ''
  });

  // Delete purchase state & modal
  const [purchaseToDelete, setPurchaseToDelete] = useState<Batch | null>(null);
  const [bulkPurchasesToDelete, setBulkPurchasesToDelete] = useState<Batch[] | null>(null);

  // Delete sale state & modal
  const [saleToDelete, setSaleToDelete] = useState<Transaction | null>(null);
  const [isDeletingMovement, setIsDeletingMovement] = useState(false);

  const [selectedPurchases, setSelectedPurchases] = useState<string[]>([]);
  const [purchaseItems, setPurchaseItems] = useState<{ productId: string; quantity: number; cost: number; productName: string }[]>([]);
  const [saleItems, setSaleItems] = useState<{ productId: string; quantity: number; salePrice: number; productName: string }[]>([]);
  const [exitItems, setExitItems] = useState<{ productId: string; quantity: number; productName: string }[]>([]);
  const [confirmationSimulation, setConfirmationSimulation] = useState<TransactionSimulationResult | null>(null);
  const [pendingTransactionData, setPendingTransactionData] = useState<{
    type: 'EXIT' | 'SALE';
    items: Array<{ productId: string; quantity: number; salePrice?: number; productName: string; description?: string }>;
    notes?: string;
    movementDate?: string;
  } | null>(null);
  const [isExecutingTransaction, setIsExecutingTransaction] = useState(false);
  const [isConsolidating, setIsConsolidating] = useState(false);

  const [notifications, setNotifications] = useState<AppNotification[]>(() => {
    const saved = localStorage.getItem('sied_notifications');
    return saved ? JSON.parse(saved) : [];
  });
  const [showNotifications, setShowNotifications] = useState(false);

  const [isRegistering, setIsRegistering] = useState(false);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // --- FIREBASE AUTH & PERSISTENCE ---
  useEffect(() => {
    const unsub = authService.onAuthChanged(async (firebaseUser) => {
      if (firebaseUser) {
        const email = firebaseUser.email || '';
        let profile = await dbService.getUserProfile(firebaseUser.uid);

        // Fetch all existing companies in Firestore to enable workspace switching & auto-discovery
        const allCompanies = await dbService.listCompaniesWithStats();
        setAvailableCompanies(allCompanies);

        let matchingTeam = null;
        if (!profile || profile.type !== 'COMPANY') {
          matchingTeam = email ? await dbService.findTeamMemberByEmail(email) : null;
          if (matchingTeam && matchingTeam.member.status === 'inactive') {
            alert('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
            await authService.logout();
            return;
          }
        }

        // Find primary company with real saved data (e.g. "prueba" or "destilera")
        const primaryCompanyWithData = allCompanies.find(c => c.productCount > 0);

        if (!profile) {
          const companyId = matchingTeam 
            ? matchingTeam.companyId 
            : (primaryCompanyWithData ? primaryCompanyWithData.id : firebaseUser.uid);
          const type: AccountType = matchingTeam ? 'USER' : 'COMPANY';
          const role = matchingTeam ? matchingTeam.member.role : 'ADMIN';
          const areas = matchingTeam ? matchingTeam.member.areas : ['INVENTARIO', 'PRODUCCION', 'REPORTES', 'CONFIGURACION'];
          
          let companyName = primaryCompanyWithData ? primaryCompanyWithData.name : 'Mi Empresa';
          if (matchingTeam) {
            const comp = await dbService.getCompany(matchingTeam.companyId);
            companyName = comp?.name || companyName;
          }

          profile = {
            id: firebaseUser.uid,
            type: type,
            name: firebaseUser.displayName || email.split('@')[0] || 'Administrador',
            email: email,
            companyName: companyName,
            linkedCompanyId: companyId,
            role: role,
            areas: areas
          };
          await dbService.setUserProfile(profile);
          if (type === 'COMPANY') {
            await dbService.createCompany(firebaseUser.uid, companyName, firebaseUser.uid);
          }
        } else {
          // If profile exists, check if user's linkedCompanyId is empty or points to an empty company
          // while existing companies have saved data
          const currentLinkedId = profile.linkedCompanyId;
          const currentHasProducts = currentLinkedId ? await dbService.hasCompanyProducts(currentLinkedId) : false;

          if ((!currentLinkedId || !currentHasProducts) && primaryCompanyWithData) {
            console.log('Connecting user to existing saved company:', primaryCompanyWithData.name, primaryCompanyWithData.id);
            profile = {
              ...profile,
              linkedCompanyId: primaryCompanyWithData.id,
              companyName: primaryCompanyWithData.name
            };
            await dbService.setUserProfile(profile);
          } else if (profile.type === 'COMPANY') {
            if (!profile.linkedCompanyId) {
              profile = { ...profile, linkedCompanyId: profile.id };
              await dbService.setUserProfile(profile);
            }
          } else if (profile.type === 'USER') {
            const currentEmail = firebaseUser.email || profile.email;
            const currentTeam = await dbService.findTeamMemberByEmail(currentEmail);
            if (currentTeam) {
              const comp = await dbService.getCompany(currentTeam.companyId);
              const companyName = comp?.name || profile.companyName;
              
              if (
                profile.linkedCompanyId !== currentTeam.companyId ||
                profile.role !== currentTeam.member.role ||
                profile.companyName !== companyName
              ) {
                profile = {
                  ...profile,
                  linkedCompanyId: currentTeam.companyId,
                  role: currentTeam.member.role,
                  areas: currentTeam.member.areas,
                  companyName: companyName
                };
                await dbService.setUserProfile(profile);
              }
            }
          }
        }

        setCurrentUser(profile);
        localStorage.setItem('sied_current_user', JSON.stringify(profile));
        localStorage.removeItem('sied_explicit_logout');
      } else {
        const isExplicitLogout = localStorage.getItem('sied_explicit_logout') === 'true';
        if (isExplicitLogout) {
          setCurrentUser(null);
          localStorage.removeItem('sied_current_user');
        } else {
          const saved = localStorage.getItem('sied_current_user');
          if (!saved) {
            setCurrentUser(null);
          }
        }
      }
    });
    return () => unsub();
  }, []);

  const handleSwitchCompany = async (companyId: string, companyName: string) => {
    if (!currentUser) return;
    const updatedProfile: Account = {
      ...currentUser,
      linkedCompanyId: companyId,
      companyName: companyName
    };
    setCurrentUser(updatedProfile);
    localStorage.setItem('sied_current_user', JSON.stringify(updatedProfile));
    await dbService.setUserProfile(updatedProfile);
  };

  // --- FIREBASE DATA LISTENERS (REAL-TIME ALL COLLECTIONS) ---
  useEffect(() => {
    if (!effectiveCompanyId) {
      setProducts([]);
      setBatches([]);
      setTransactions([]);
      setProductionOrders([]);
      setTeam([]);
      setAuditLogs([]);
      setNotifications([]);
      return;
    }

    const unsubs = [
      dbService.subscribeToCollection<Product>(effectiveCompanyId, 'products', setProducts),
      dbService.subscribeToCollection<Batch>(effectiveCompanyId, 'batches', setBatches),
      dbService.subscribeToCollection<Transaction>(effectiveCompanyId, 'transactions', setTransactions),
      dbService.subscribeToCollection<ProductionOrderRecord>(effectiveCompanyId, 'productionOrders', setProductionOrders),
      dbService.subscribeToCollection<TeamMember>(effectiveCompanyId, 'team', setTeam),
      dbService.subscribeToCollection<AuditEntry>(effectiveCompanyId, 'auditLogs', setAuditLogs),
      dbService.subscribeToCollection<AppNotification>(effectiveCompanyId, 'notifications', setNotifications),
    ];

    return () => unsubs.forEach(unsub => unsub());
  }, [effectiveCompanyId]);

  // --- REAL-TIME TEAM MEMBER STATUS & ACCESS SYNC ---
  useEffect(() => {
    if (!currentUser || currentUser.type !== 'USER' || team.length === 0) return;

    const myEmail = currentUser.email?.trim().toLowerCase();
    if (!myEmail) return;

    const myMember = team.find(m => m.email.trim().toLowerCase() === myEmail);
    if (myMember) {
      // 1. Check if inactive
      if (myMember.status === 'inactive') {
        alert('Su usuario ha sido desactivado por el administrador. Se cerrará la sesión de forma inmediata.');
        authService.logout();
        return;
      }

      // 2. Check if roles or areas changed
      const arraysEqual = (a: string[], b: string[]) => {
        if (a.length !== b.length) return false;
        const sortedA = [...a].sort();
        const sortedB = [...b].sort();
        return sortedA.every((val, index) => val === sortedB[index]);
      };

      const currentAreas = currentUser.areas || [];
      const currentRole = currentUser.role || 'OPERATOR';

      if (myMember.role !== currentRole || !arraysEqual(myMember.areas || [], currentAreas)) {
        const updatedProfile = {
          ...currentUser,
          role: myMember.role,
          areas: myMember.areas || []
        };
        setCurrentUser(updatedProfile);
        localStorage.setItem('sied_current_user', JSON.stringify(updatedProfile));
        dbService.setUserProfile(updatedProfile);
      }
    }
  }, [team, currentUser]);

  const addNotification = async (notif: Omit<AppNotification, 'id' | 'timestamp' | 'read'>) => {
    if (!effectiveCompanyId) return;
    const newNotif: AppNotification = {
      ...notif,
      id: Math.random().toString(36).substr(2, 9),
      timestamp: new Date().toISOString(),
      read: false
    };
    await dbService.saveNotification(effectiveCompanyId, newNotif);
    addToast(notif);
  };

  const markAsRead = async (id: string) => {
    if (!effectiveCompanyId) return;
    const notif = notifications.find(n => n.id === id);
    if (notif) {
      await dbService.saveNotification(effectiveCompanyId, { ...notif, read: true });
    }
  };

  const clearNotifications = async () => {
    if (!effectiveCompanyId) return;
    for (const n of notifications) {
      await dbService.deleteItem(effectiveCompanyId, 'notifications', n.id);
    }
  };

  const handleSeedSystemData = async () => {
    if (!effectiveCompanyId) {
      alert('No se detectó un identificador de empresa válido.');
      return;
    }
    const confirmSeed = window.confirm(
      '¿Desea cargar los productos y datos de demostración de "Destilería Ormón" en esta empresa?\n\n' +
      'NOTA: Sus productos y lotes existentes NO serán eliminados, pero se agregarán los productos de demostración al inventario.'
    );
    if (!confirmSeed) return;
    setIsSeedingData(true);
    try {
      await dbService.seedDefaultCompanyData(
        effectiveCompanyId,
        currentUser?.companyName || 'DESTILERÍA ORMÓN',
        currentUser?.name || 'Administrador'
      );
      addNotification({
        type: 'SUCCESS',
        title: 'Base de Datos Sincronizada',
        message: 'Se cargaron todas las materias primas, insumos, productos terminados (recetas y empaques) y movimientos al Kardex.',
        linkTab: 'inventory'
      });
      alert('¡Base de datos cargada con éxito! Se sincronizaron los productos, inventario, recetas y movimientos de Kardex.');
    } catch (err) {
      console.error('Error al sincronizar datos del sistema:', err);
      alert('Error al poblar la base de datos: ' + (err instanceof Error ? err.message : String(err)));
    } finally {
      setIsSeedingData(false);
    }
  };

  // Kardex Filters
  const [kardexFilterProduct, setKardexFilterProduct] = useState('');
  const [kardexFilterType, setKardexFilterType] = useState<'ALL' | 'ENTRY' | 'EXIT' | 'PURCHASE' | 'SALE'>('ALL');
  const [kardexFilterStartDate, setKardexFilterStartDate] = useState('');
  const [kardexFilterEndDate, setKardexFilterEndDate] = useState('');
  const [kardexSearch, setKardexSearch] = useState('');
  const [inventorySearch, setInventorySearch] = useState('');
  const [auditSearch, setAuditSearch] = useState('');
  const [auditFilterAction, setAuditFilterAction] = useState<AuditEntry['action'] | 'ALL'>('ALL');
  const [auditFilterEntity, setAuditFilterEntity] = useState<AuditEntry['entityType'] | 'ALL'>('ALL');
  const [auditFilterStartDate, setAuditFilterStartDate] = useState('');
  const [auditFilterEndDate, setAuditFilterEndDate] = useState('');
  const [peopleSearch, setPeopleSearch] = useState('');
  const [peopleRoleFilter, setPeopleRoleFilter] = useState<'ALL' | 'ADMIN' | 'OPERATOR'>('ALL');
  const [kardexCurrentPage, setKardexCurrentPage] = useState(1);
  const KARDEX_ITEMS_PER_PAGE = 50;

  // Purchase Filters
  const [purchaseFilterSupplier, setPurchaseFilterSupplier] = useState('');
  const [purchaseFilterStartDate, setPurchaseFilterStartDate] = useState('');
  const [purchaseFilterEndDate, setPurchaseFilterEndDate] = useState('');

  const filteredPurchaseReports = useMemo(() => {
    // Purchases are entries in batches with type 'PURCHASE'
    return batches
      .filter(b => b.type === 'PURCHASE')
      .filter(b => {
        const product = products.find(p => p.id === b.productId);
        const supplierMatch = b.supplier.toLowerCase().includes(purchaseFilterSupplier.toLowerCase()) || 
                             (product?.name || '').toLowerCase().includes(purchaseFilterSupplier.toLowerCase());
        
        let dateMatch = true;
        if (purchaseFilterStartDate) {
          dateMatch = dateMatch && new Date(b.entryDate) >= new Date(purchaseFilterStartDate);
        }
        if (purchaseFilterEndDate) {
          const end = new Date(purchaseFilterEndDate);
          end.setHours(23, 59, 59, 999);
          dateMatch = dateMatch && new Date(b.entryDate) <= end;
        }

        return supplierMatch && dateMatch;
      })
      .sort((a, b) => new Date(b.entryDate).getTime() - new Date(a.entryDate).getTime());
  }, [batches, products, purchaseFilterSupplier, purchaseFilterStartDate, purchaseFilterEndDate]);

  // Sales Filters & Metrics
  const [saleFilterSearch, setSaleFilterSearch] = useState('');
  const [saleFilterStartDate, setSaleFilterStartDate] = useState('');
  const [saleFilterEndDate, setSaleFilterEndDate] = useState('');

  const filteredSaleReports = useMemo(() => {
    return transactions
      .filter(t => t.type === 'SALE')
      .filter(t => {
        const product = products.find(p => p.id === t.productId);
        const searchLower = saleFilterSearch.toLowerCase();
        const matchesSearch = !saleFilterSearch ||
          (product?.name || '').toLowerCase().includes(searchLower) ||
          (product?.category || '').toLowerCase().includes(searchLower) ||
          (t.reference || '').toLowerCase().includes(searchLower) ||
          (t.description || '').toLowerCase().includes(searchLower) ||
          (t.performedBy || '').toLowerCase().includes(searchLower);

        let dateMatch = true;
        if (saleFilterStartDate) {
          dateMatch = dateMatch && new Date(t.date) >= new Date(saleFilterStartDate);
        }
        if (saleFilterEndDate) {
          const end = new Date(saleFilterEndDate);
          end.setHours(23, 59, 59, 999);
          dateMatch = dateMatch && new Date(t.date) <= end;
        }

        return matchesSearch && dateMatch;
      })
      .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
  }, [transactions, products, saleFilterSearch, saleFilterStartDate, saleFilterEndDate]);

  const salesSummary = useMemo(() => {
    let totalRevenue = 0;
    let totalCogs = 0;
    let totalUnits = 0;
    filteredSaleReports.forEach(s => {
      const unitSalePrice = s.salePrice || s.unitCost || 0;
      totalRevenue += s.quantity * unitSalePrice;
      totalCogs += s.quantity * s.unitCost;
      totalUnits += s.quantity;
    });
    const grossProfit = totalRevenue - totalCogs;
    const profitMargin = totalRevenue > 0 ? (grossProfit / totalRevenue) * 100 : 0;
    return { totalRevenue, totalCogs, totalUnits, grossProfit, profitMargin };
  }, [filteredSaleReports]);

  const filteredTeam = useMemo(() => {
    return team
      .filter(m => {
        const searchMatch = m.name.toLowerCase().includes(peopleSearch.toLowerCase()) || 
                           m.email.toLowerCase().includes(peopleSearch.toLowerCase());
        const roleMatch = peopleRoleFilter === 'ALL' || m.role === peopleRoleFilter;
        return searchMatch && roleMatch;
      });
  }, [team, peopleSearch, peopleRoleFilter]);

  const filteredAuditLogs = useMemo(() => {
    return auditLogs.filter(log => {
      const s = auditSearch.toLowerCase();
      const matchesSearch = !auditSearch || 
                          log.details.toLowerCase().includes(s) || 
                          log.userName.toLowerCase().includes(s) || 
                          log.entityId.toLowerCase().includes(s);
      const matchesAction = auditFilterAction === 'ALL' || log.action === auditFilterAction;
      const matchesEntity = auditFilterEntity === 'ALL' || log.entityType === auditFilterEntity;
      
      let dateMatch = true;
      if (auditFilterStartDate) {
        dateMatch = dateMatch && new Date(log.timestamp) >= new Date(auditFilterStartDate);
      }
      if (auditFilterEndDate) {
        const end = new Date(auditFilterEndDate);
        end.setHours(23, 59, 59, 999);
        dateMatch = dateMatch && new Date(log.timestamp) <= end;
      }

      return matchesSearch && matchesAction && matchesEntity && dateMatch;
    });
  }, [auditLogs, auditSearch, auditFilterAction, auditFilterEntity, auditFilterStartDate, auditFilterEndDate]);

  const auditStats = useMemo(() => {
    return {
      total: filteredAuditLogs.length,
      creations: filteredAuditLogs.filter(l => l.action === 'CREATE').length,
      updates: filteredAuditLogs.filter(l => l.action === 'UPDATE').length,
      deletions: filteredAuditLogs.filter(l => l.action === 'DELETE').length,
    };
  }, [filteredAuditLogs]);

  // Reset page when filters change
  useEffect(() => {
    setKardexCurrentPage(1);
  }, [kardexFilterProduct, kardexFilterType, kardexFilterStartDate, kardexFilterEndDate, kardexSearch]);

  const [formData, setFormData] = useState({ 
    quantity: 0, 
    cost: 0, 
    minStock: 0,
    supplier: '', 
    name: '', 
    description: '', 
    category: '', 
    unit: 'litro' as const,
    type: 'materia prima' as const,
    productDescription: '', // description of entry/exit
    entryDate: new Date().toISOString().split('T')[0]
  });

  const addAuditLog = (action: AuditEntry['action'], entityType: AuditEntry['entityType'], entityId: string, details: string, prev?: any, curr?: any) => {
    if (!currentUser) return;
    const companyId = currentUser.linkedCompanyId || currentUser.id;
    
    const log: AuditEntry = {
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      userId: currentUser.id,
      userName: currentUser.name,
      action,
      entityType,
      entityId,
      details,
      previousValue: prev ? JSON.stringify(prev) : undefined,
      newValue: curr ? JSON.stringify(curr) : undefined
    };

    dbService.addAuditLog(companyId, log);
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isRegistering) return;
    setIsRegistering(true);
    
    try {
      const type = registerForm === 'empresa' ? 'COMPANY' : 'USER';
      const companyName = registerForm === 'empresa' ? regFormData.company : 'Personal';
      const name = registerForm === 'empresa' ? regFormData.company : regFormData.name;
      
      console.log('Attempting registration for:', regFormData.email);
      
      const profile = await authService.register(
        regFormData.email, 
        regFormData.password, 
        name, 
        type, 
        companyName
      );

      console.log('Registration successful, profile created:', profile.id);

      // If it's a USER, we need to check if they're already part of a company
      if (type === 'USER') {
        const linkedProfile = await dbService.getUserProfile(profile.id);
        if (linkedProfile) {
          setCurrentUser(linkedProfile);
        } else {
          setCurrentUser(profile);
        }
      } else {
        setCurrentUser(profile);
      }

      setRegisterForm(null);
      setShowRegisterOptions(false);
      setRegFormData({ company: '', name: '', email: '', password: '' });
    } catch (err: any) {
      console.error('Registration error detail:', err);
      let errorMsg = 'Error al registrar.';
      const errorCode = err.code || '';
      const errorMessage = err.message || '';

      if (errorCode === 'auth/operation-not-allowed' || errorMessage.includes('auth/operation-not-allowed')) {
        errorMsg = 'ERROR CONFIGURACIÓN: El método de registro con correo está DESACTIVADO en Firebase.\n\nPara solucionar esto:\n1. Ve a: https://console.firebase.google.com/project/' + (firebaseConfig as any).projectId + '/authentication/providers\n2. Habilita "Email/Password"\n3. Guarda los cambios.\n\nAlternativamente, usa "Continuar con Google".';
      } else if (errorCode === 'auth/email-already-in-use' || errorMessage.includes('auth/email-already-in-use')) {
        errorMsg = 'Este correo ya está registrado.';
      } else if (errorCode === 'auth/weak-password' || errorMessage.includes('auth/weak-password')) {
        errorMsg = 'La contraseña es muy débil (mínimo 6 caracteres).';
      } else if (errorCode === 'auth/invalid-email' || errorMessage.includes('auth/invalid-email')) {
        errorMsg = 'Correo electrónico inválido.';
      } else if (err instanceof Error) {
        errorMsg = err.message;
      }
      alert(errorMsg);
    } finally {
      setIsRegistering(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isLoggingIn) return;
    setIsLoggingIn(true);
    
    try {
      localStorage.removeItem('sied_explicit_logout');
      const profile = await authService.login(loginFormData.email, loginFormData.password);
      if (profile) {
        setCurrentUser(profile);
        localStorage.setItem('sied_current_user', JSON.stringify(profile));
        setLoginFormData({ email: '', password: '' });
      } else {
        alert('No se encontró el perfil de usuario en el servidor.');
      }
    } catch (err: any) {
      console.error('Login error detail:', err);
      let errorMsg = 'Error en el inicio de sesión. Verifique sus credenciales.';
      const errorCode = err.code || '';
      const errorMessage = err.message || '';

      if (errorMessage.includes('contraseña anterior ha sido eliminada') || errorMessage.includes('inactivo') || errorMessage.includes('únicamente con la nueva contraseña')) {
        errorMsg = errorMessage;
      } else if (errorCode === 'auth/operation-not-allowed' || errorMessage.includes('auth/operation-not-allowed')) {
        errorMsg = 'ERROR: El inicio de sesión con correo está DESACTIVADO. Actívalo en la Consola de Firebase: https://console.firebase.google.com/project/' + (firebaseConfig as any).projectId + '/authentication/providers';
      } else if (errorCode === 'auth/user-not-found' || errorMessage.includes('auth/user-not-found')) {
        errorMsg = 'Usuario no encontrado.';
      } else if (errorCode === 'auth/wrong-password' || errorMessage.includes('auth/wrong-password')) {
        errorMsg = 'Contraseña incorrecta.';
      } else if (errorCode === 'auth/invalid-credential' || errorMessage.includes('auth/invalid-credential')) {
        errorMsg = 'Credenciales inválidas.';
      } else if (errorMessage) {
        errorMsg = errorMessage;
      }
      alert(errorMsg);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleResetPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setResetError('');

    const targetEmail = resetEmail.trim().toLowerCase();
    if (!targetEmail) {
      setResetError('Por favor ingresa el correo electrónico asociado a la cuenta.');
      return;
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(targetEmail)) {
      setResetError('Por favor introduce un formato de correo electrónico válido.');
      return;
    }

    if (newPassword.length < 6) {
      setResetError('La contraseña debe tener al menos 6 caracteres.');
      return;
    }

    if (newPassword !== confirmPassword) {
      setResetError('Las contraseñas no coinciden. Por favor verifica ambos campos.');
      return;
    }

    setIsResettingPassword(true);
    try {
      await authService.resetOrUpdatePassword(newPassword, resetOobCode || undefined, targetEmail);
      setResetSuccess(true);
      // Pre-fill login credentials with the target email and new password so login is immediate
      setLoginFormData({ email: targetEmail, password: newPassword });
      
      setTimeout(() => {
        setIsForgotPassword(false);
        setResetSuccess(false);
        setResetEmail('');
        setNewPassword('');
        setConfirmPassword('');
      }, 1800);
    } catch (err: any) {
      console.error('Password reset error:', err);
      let msg = 'Error al cambiar la contraseña.';
      if (err.code === 'auth/weak-password') {
        msg = 'La contraseña es muy débil (mínimo 6 caracteres).';
      } else if (err.code === 'auth/invalid-action-code') {
        msg = 'El código o enlace de recuperación ha expirado o ya fue utilizado.';
      } else if (err.message) {
        msg = err.message;
      }
      setResetError(msg);
    } finally {
      setIsResettingPassword(false);
    }
  };

  const handleLogout = async () => {
    localStorage.setItem('sied_explicit_logout', 'true');
    await authService.logout();
    setCurrentUser(null);
    localStorage.removeItem('sied_current_user');
    setActiveTab('dashboard');
  };

  const accountType = currentUser?.type;
  const companyName = currentUser?.companyName;
  const role = currentUser?.role || (accountType === 'COMPANY' ? 'ADMIN' : 'OPERATOR');
  const userAreas = currentUser?.areas || [];

  // Derived stock
  const productStock = useMemo(() => {
    const stock: Record<string, number> = {};
    products.forEach(p => {
      stock[p.id] = batches
        .filter(b => b.productId === p.id)
        .reduce((sum, b) => sum + b.remainingQuantity, 0);
    });
    return stock;
  }, [products, batches]);

  // Low Stock Monitor with ref tracking to prevent infinite notification loops
  const notifiedLowStockRef = useRef<Set<string>>(new Set());

  useEffect(() => {
    if (!currentUser?.linkedCompanyId || products.length === 0) return;

    products.forEach(product => {
      const stock = productStock[product.id] ?? 0;
      if (stock <= product.minStock) {
        // Check if we already have an unread notification for this product or already notified
        const hasUnread = notifications.some(n => n.productId === product.id && n.type === 'LOW_STOCK' && !n.read);
        const alreadyNotified = notifiedLowStockRef.current.has(product.id);
        
        if (!hasUnread && !alreadyNotified) {
          notifiedLowStockRef.current.add(product.id);
          if (stock === 0) {
            addNotification({
              type: 'LOW_STOCK',
              title: 'Stock Agotado',
              message: `¡ALERTA! El producto ${product.name} se ha agotado completamente.`,
              linkTab: 'inventory',
              productId: product.id
            });
          } else {
            addNotification({
              type: 'LOW_STOCK',
              title: 'Stock Crítico',
              message: `El producto ${product.name} está por debajo del mínimo. Stock: ${stock} / Mín: ${product.minStock}`,
              linkTab: 'inventory',
              productId: product.id
            });
          }
        }
      } else {
        // Stock recovered above threshold, allow alert next time it drops
        notifiedLowStockRef.current.delete(product.id);
      }
    });
  }, [productStock, products, notifications, currentUser?.linkedCompanyId]);
  
  const filteredKardex = useMemo(() => {
    // Collect all finished products IDs
    const finishedProductIds = new Set(
      products.filter(p => p.type === 'producto terminado').map(p => p.id)
    );

    // Only process kardex for finished products
    const finishedProductsList = products.filter(p => p.type === 'producto terminado');
    const fullKardex = generateUnifiedKardex(finishedProductsList, batches, transactions);

    return fullKardex.filter(entry => {
      // Must be a finished product
      if (!entry.productId || !finishedProductIds.has(entry.productId)) {
        return false;
      }

      // Search Box Filter
      const searchLower = kardexSearch.toLowerCase();
      const matchesSearch = !kardexSearch || 
                           entry.productName?.toLowerCase().includes(searchLower) ||
                           entry.reference?.toLowerCase().includes(searchLower) ||
                           entry.performedBy?.toLowerCase().includes(searchLower);

      // Product Name Filter (dropdown or search)
      const matchesProduct = !kardexFilterProduct || 
                            kardexFilterProduct === 'ALL_FINISHED' || 
                            entry.productId === kardexFilterProduct ||
                            entry.productName?.toLowerCase().includes(kardexFilterProduct.toLowerCase());
      
      // Type Filter
      const matchesType = kardexFilterType === 'ALL' || entry.type === kardexFilterType;
      
      // Date Filter
      const entryDate = new Date(entry.date);
      const matchesStart = !kardexFilterStartDate || entryDate >= new Date(kardexFilterStartDate);
      
      // For end date, we add 1 day or check if it's inclusive
      let matchesEnd = true;
      if (kardexFilterEndDate) {
        const endDate = new Date(kardexFilterEndDate);
        endDate.setHours(23, 59, 59, 999);
        matchesEnd = entryDate <= endDate;
      }
      
      return matchesSearch && matchesProduct && matchesType && matchesStart && matchesEnd;
    });
  }, [products, batches, transactions, kardexFilterProduct, kardexFilterType, kardexFilterStartDate, kardexFilterEndDate, kardexSearch]);

  const paginatedKardex = useMemo(() => {
    const start = (kardexCurrentPage - 1) * KARDEX_ITEMS_PER_PAGE;
    return filteredKardex.slice(start, start + KARDEX_ITEMS_PER_PAGE);
  }, [filteredKardex, kardexCurrentPage]);

  const totalKardexPages = Math.ceil(filteredKardex.length / KARDEX_ITEMS_PER_PAGE);

  const dashboardStats = useMemo(() => {
    const cogs = transactions.filter(t => t.type === 'SALE').reduce((acc, t) => acc + (t.quantity * t.unitCost), 0);
    const mermas = transactions.filter(t => t.type === 'EXIT').reduce((acc, t) => acc + (t.quantity * t.unitCost), 0);
    const valuation = batches.reduce((acc, b) => acc + (b.remainingQuantity * b.unitCost), 0);
    return { cogs, mermas, valuation };
  }, [transactions, batches]);

  const salesTrendData = useMemo(() => {
    const rawSalesByDate: Record<string, number> = {};
    
    transactions
      .filter(t => t.type === 'SALE')
      .forEach(t => {
        const dateKey = t.date.split('T')[0];
        rawSalesByDate[dateKey] = (rawSalesByDate[dateKey] || 0) + (t.quantity * t.unitCost);
      });

    return Object.entries(rawSalesByDate)
      .map(([date, amount]) => ({ date, amount }))
      .sort((a, b) => a.date.localeCompare(b.date))
      .map(item => ({
        ...item,
        formattedDate: format(new Date(item.date + 'T12:00:00'), 'dd MMM', { locale: es })
      }));
  }, [transactions]);

  const handleAddProduct = (e: React.FormEvent) => {
    e.preventDefault();
    const newProduct: Product = {
      id: crypto.randomUUID(),
      name: formData.name,
      description: formData.description,
      category: formData.category,
      unit: formData.unit as any,
      type: formData.type as any,
      minStock: formData.minStock,
      recipe: (formData.type === 'producto terminado' || formData.type === 'producción') && productRecipe.length > 0 
        ? productRecipe 
        : undefined,
      packaging: (formData.type === 'producto terminado' || formData.type === 'producción') && productPackaging.packagingType !== 'ninguno'
        ? productPackaging
        : undefined
    };
    if (effectiveCompanyId) {
      dbService.saveProduct(effectiveCompanyId, newProduct);
    }
    addAuditLog('CREATE', 'PRODUCT', newProduct.id, `Producto creado: ${newProduct.name} (${newProduct.type})`, null, newProduct);
    addNotification({
      type: 'SUCCESS',
      title: 'Producto Registrado',
      message: `${newProduct.name} ha sido añadido al inventario.`,
      linkTab: 'inventory'
    });
    setShowModal(null);
    setProductRecipe([]);
    setProductPackaging({ packagingType: 'ninguno', unitsPerBox: 0, unitsPerBundle: 0 });
    setFormData({ 
      quantity: 0, 
      cost: 0, 
      minStock: 0, 
      supplier: '', 
      name: '', 
      description: '', 
      category: '', 
      unit: 'litro',
      type: 'materia prima',
      productDescription: '',
      entryDate: new Date().toISOString().split('T')[0]
    });
  };

  const handleAddEntry = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedProductId) return;

    const nowTime = new Date().toTimeString().split(' ')[0];
    const resolvedEntryDate = formData.entryDate
      ? (formData.entryDate.includes('T') ? formData.entryDate : `${formData.entryDate}T${nowTime}`)
      : new Date().toISOString();

    const newBatch: Batch = {
      id: crypto.randomUUID(),
      productId: selectedProductId,
      entryDate: resolvedEntryDate,
      quantity: formData.quantity,
      remainingQuantity: formData.quantity,
      unitCost: formData.cost,
      supplier: formData.supplier || 'N/A',
      description: formData.productDescription,
      type: 'ENTRY',
      performedBy: currentUser?.name || currentUser?.email || 'Sistema'
    };

    if (effectiveCompanyId) {
      dbService.saveBatch(effectiveCompanyId, newBatch);
    }
    addAuditLog('CREATE', 'BATCH', newBatch.id, `Entrada manual (${format(new Date(resolvedEntryDate), 'dd/MM/yyyy')}): ${newBatch.quantity} unidades para ${products.find(p => p.id === selectedProductId)?.name}`, null, newBatch);
    setShowModal(null);
    setFormData({ 
      quantity: 0, 
      cost: 0, 
      minStock: 0,
      supplier: '', 
      name: '', 
      description: '', 
      category: '', 
      unit: 'litro',
      type: 'materia prima',
      productDescription: '',
      entryDate: new Date().toISOString().split('T')[0]
    });
  };

  const addItemToExit = () => {
    if (!selectedProductId || formData.quantity <= 0) return;
    const currentStock = productStock[selectedProductId] || 0;
    const alreadyInList = exitItems.filter(i => i.productId === selectedProductId).reduce((acc, i) => acc + i.quantity, 0);
    if (formData.quantity + alreadyInList > currentStock) {
      alert(`Stock insuficiente. Stock actual: ${currentStock}, ya en lista: ${alreadyInList}`);
      return;
    }
    const productName = products.find(p => p.id === selectedProductId)?.name || '';
    setExitItems(prev => [
      ...prev,
      {
        productId: selectedProductId,
        quantity: formData.quantity,
        productName
      }
    ]);
    setFormData({ ...formData, quantity: 0 });
    setSelectedProductId(null);
  };

  const addItemToSale = () => {
    if (!selectedProductId || formData.quantity <= 0) return;
    const currentStock = productStock[selectedProductId] || 0;
    const alreadyInList = saleItems.filter(i => i.productId === selectedProductId).reduce((acc, i) => acc + i.quantity, 0);
    if (formData.quantity + alreadyInList > currentStock) {
      alert(`Stock insuficiente. Stock actual: ${currentStock}, ya en venta: ${alreadyInList}`);
      return;
    }
    const productName = products.find(p => p.id === selectedProductId)?.name || '';
    setSaleItems(prev => [
      ...prev,
      {
        productId: selectedProductId,
        quantity: formData.quantity,
        salePrice: formData.cost,
        productName
      }
    ]);
    setFormData({ ...formData, quantity: 0, cost: 0 });
    setSelectedProductId(null);
  };

  const handleInitiateExit = (e: React.FormEvent) => {
    e.preventDefault();
    const itemsToProcess = [...exitItems];
    if (selectedProductId && formData.quantity > 0) {
      const productName = products.find(p => p.id === selectedProductId)?.name || '';
      itemsToProcess.push({
        productId: selectedProductId,
        quantity: formData.quantity,
        productName
      });
    }

    if (itemsToProcess.length === 0) {
      alert('Debe especificar al menos un producto y una cantidad válida mayor a cero.');
      return;
    }

    const simulation = simulateExitImpact(itemsToProcess, batches, products, 'EXIT');
    if (simulation.hasError) {
      const firstErr = simulation.items.find(i => i.error)?.error;
      alert(firstErr || 'Stock insuficiente para realizar la salida solicitada.');
      return;
    }

    const nowTime = new Date().toTimeString().split(' ')[0];
    const resolvedMovementDate = formData.entryDate 
      ? (formData.entryDate.includes('T') ? formData.entryDate : `${formData.entryDate}T${nowTime}`)
      : new Date().toISOString();

    setPendingTransactionData({
      type: 'EXIT',
      items: itemsToProcess.map(i => ({ ...i, description: formData.productDescription })),
      notes: formData.productDescription,
      movementDate: resolvedMovementDate
    });
    setConfirmationSimulation(simulation);
    setShowModal(null);
  };

  const handleInitiateSale = (e: React.FormEvent) => {
    e.preventDefault();
    const itemsToProcess = [...saleItems];
    if (selectedProductId && formData.quantity > 0) {
      const productName = products.find(p => p.id === selectedProductId)?.name || '';
      itemsToProcess.push({
        productId: selectedProductId,
        quantity: formData.quantity,
        salePrice: formData.cost,
        productName
      });
    }

    if (itemsToProcess.length === 0) {
      alert('Debe especificar al menos un producto y una cantidad válida mayor a cero.');
      return;
    }

    const simulation = simulateExitImpact(itemsToProcess, batches, products, 'SALE');
    if (simulation.hasError) {
      const firstErr = simulation.items.find(i => i.error)?.error;
      alert(firstErr || 'Stock insuficiente para realizar la venta solicitada.');
      return;
    }

    const nowTime = new Date().toTimeString().split(' ')[0];
    const resolvedMovementDate = formData.entryDate 
      ? (formData.entryDate.includes('T') ? formData.entryDate : `${formData.entryDate}T${nowTime}`)
      : new Date().toISOString();

    setPendingTransactionData({
      type: 'SALE',
      items: itemsToProcess.map(i => ({ ...i, description: formData.productDescription })),
      notes: formData.productDescription,
      movementDate: resolvedMovementDate
    });
    setConfirmationSimulation(simulation);
    setShowModal(null);
  };

  const handleConfirmAndExecuteTransaction = async () => {
    if (!pendingTransactionData || isExecutingTransaction) return;
    setIsExecutingTransaction(true);

    try {
      let currentBatches = [...batches];
      const allNewTransactions: Transaction[] = [];

      for (const item of pendingTransactionData.items) {
        const { updatedBatches, transactions: newTx } = processExit(
          item.quantity,
          currentBatches,
          item.productId,
          pendingTransactionData.notes || item.description,
          pendingTransactionData.type,
          currentUser?.name || currentUser?.email || 'Sistema',
          item.salePrice,
          pendingTransactionData.movementDate
        );

        currentBatches = updatedBatches;
        allNewTransactions.push(...newTx);
      }

      if (effectiveCompanyId) {
        // Save all updated batches and new transactions
        currentBatches
          .filter(b => {
            const oldB = batches.find(ob => ob.id === b.id);
            return oldB && oldB.remainingQuantity !== b.remainingQuantity;
          })
          .forEach(b => dbService.saveBatch(effectiveCompanyId, b));

        allNewTransactions.forEach(tx => dbService.saveTransaction(effectiveCompanyId, tx));
      }

      const isSale = pendingTransactionData.type === 'SALE';
      const itemsCount = pendingTransactionData.items.length;
      const totalUnits = pendingTransactionData.items.reduce((acc, i) => acc + i.quantity, 0);

      addNotification({
        type: isSale ? 'SUCCESS' : 'WARNING',
        title: isSale 
          ? (itemsCount > 1 ? 'Venta Masiva Registrada' : 'Venta Registrada') 
          : (itemsCount > 1 ? 'Salida Masiva Registrada' : 'Ajuste de Salida'),
        message: isSale
          ? `Se han vendido ${totalUnits} unidades (${itemsCount} productos) exitosamente.`
          : `Se ha registrado una salida de ${totalUnits} unidades (${itemsCount} productos).`,
        linkTab: isSale ? 'sales' : 'inventory'
      });

      allNewTransactions.forEach(tx => {
        const prod = products.find(p => p.id === tx.productId);
        addAuditLog(
          'CREATE',
          'TRANSACTION',
          tx.id,
          `${isSale ? 'Venta' : 'Salida'}: ${tx.quantity} unidades de ${prod?.name || 'Producto'} (PEPS)`,
          null,
          tx
        );
      });

      // Clear state after success
      setConfirmationSimulation(null);
      setPendingTransactionData(null);
      setSaleItems([]);
      setExitItems([]);
      setSelectedProductId(null);
      setFormData({
        quantity: 0,
        cost: 0,
        minStock: 0,
        supplier: '',
        name: '',
        description: '',
        category: '',
        unit: 'litro',
        type: 'materia prima',
        productDescription: '',
        entryDate: new Date().toISOString().split('T')[0]
      });
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Error al ejecutar la transacción');
    } finally {
      setIsExecutingTransaction(false);
    }
  };

  const handleBackToEdit = () => {
    if (!pendingTransactionData) {
      setConfirmationSimulation(null);
      return;
    }
    const targetModal = pendingTransactionData.type === 'SALE' ? 'sale' : 'exit';
    setConfirmationSimulation(null);
    setShowModal(targetModal);
  };

  const handleCancelConfirmation = () => {
    setConfirmationSimulation(null);
    setPendingTransactionData(null);
  };

  const addItemToPurchase = () => {
    if (!selectedProductId || formData.quantity <= 0) return;
    const productName = products.find(p => p.id === selectedProductId)?.name || '';
    setPurchaseItems(prev => [
      ...prev, 
      { 
        productId: selectedProductId, 
        quantity: formData.quantity, 
        cost: formData.cost, 
        productName 
      }
    ]);
    setFormData({ ...formData, quantity: 0, cost: 0 });
    setSelectedProductId(null);
  };

  const handleAddPurchase = (e: React.FormEvent) => {
    e.preventDefault();
    
    const itemsToProcess = [...purchaseItems];
    if (selectedProductId && formData.quantity > 0) {
      const productName = products.find(p => p.id === selectedProductId)?.name || '';
      itemsToProcess.push({
        productId: selectedProductId,
        quantity: formData.quantity,
        cost: formData.cost,
        productName
      });
    }

    if (itemsToProcess.length === 0) return;

    // Consolidate items by productId
    const consolidatedMap = new Map<string, { quantity: number; totalCost: number; productName: string }>();
    
    itemsToProcess.forEach(item => {
      const existing = consolidatedMap.get(item.productId);
      if (existing) {
        consolidatedMap.set(item.productId, {
          ...existing,
          quantity: existing.quantity + item.quantity,
          totalCost: existing.totalCost + (item.cost * item.quantity)
        });
      } else {
        consolidatedMap.set(item.productId, {
          productId: item.productId,
          quantity: item.quantity,
          totalCost: item.cost * item.quantity,
          productName: item.productName
        } as any);
      }
    });

    const nowTime = new Date().toTimeString().split(' ')[0];
    const resolvedEntryDate = formData.entryDate
      ? (formData.entryDate.includes('T') ? formData.entryDate : `${formData.entryDate}T${nowTime}`)
      : new Date().toISOString();

    const newBatches: Batch[] = Array.from(consolidatedMap.entries()).map(([productId, data]: any) => ({
      id: crypto.randomUUID(),
      productId: productId,
      entryDate: resolvedEntryDate,
      quantity: data.quantity,
      remainingQuantity: data.quantity,
      unitCost: data.totalCost / (data.quantity || 1),
      supplier: formData.supplier || 'Proveedor',
      description: formData.productDescription || `Compra consolidada (${itemsToProcess.filter(i => i.productId === productId).length} registros)`,
      type: 'PURCHASE',
      performedBy: currentUser?.name || currentUser?.email || 'Sistema'
    }));

    if (effectiveCompanyId) {
      newBatches.forEach(nb => dbService.saveBatch(effectiveCompanyId, nb));
    }
    
    addNotification({
      type: 'SUCCESS',
      title: 'Compra Registrada',
      message: `Se han registrado ${newBatches.length} productos (consolidación aplicada).`,
      linkTab: 'purchases'
    });
    
    newBatches.forEach(nb => {
      const productName = products.find(p => p.id === nb.productId)?.name;
      addAuditLog('CREATE', 'BATCH', nb.id, `Compra registrada: ${nb.quantity} unidades de ${productName} a ${nb.supplier}`, null, nb);
    });

    setShowModal(null);
    setPurchaseItems([]);
    setFormData({ 
      quantity: 0, 
      cost: 0, 
      minStock: 0,
      supplier: '', 
      name: '', 
      description: '', 
      category: '', 
      unit: 'litro',
      type: 'materia prima',
      productDescription: '',
      entryDate: new Date().toISOString().split('T')[0]
    });
  };

  const handleEditBatch = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedBatchId || !effectiveCompanyId) return;

    const prevBatch = batches.find(b => b.id === selectedBatchId);
    if (prevBatch) {
        // Calculate consumed portion to maintain inventory integrity
        const consumed = Math.max(0, prevBatch.quantity - prevBatch.remainingQuantity);
        const newRemaining = Math.max(0, formData.quantity - consumed);
        const nowTime = prevBatch.entryDate.includes('T') ? prevBatch.entryDate.split('T')[1] : '12:00:00';
        const resolvedEntryDate = formData.entryDate
          ? (formData.entryDate.includes('T') ? formData.entryDate : `${formData.entryDate}T${nowTime}`)
          : prevBatch.entryDate;

        const updatedBatch = {
          ...prevBatch,
          entryDate: resolvedEntryDate,
          quantity: formData.quantity,
          remainingQuantity: newRemaining,
          unitCost: formData.cost,
          supplier: formData.supplier || 'N/A',
          description: formData.productDescription,
          performedBy: currentUser?.name || currentUser?.email || 'Sistema'
        };
        dbService.saveBatch(effectiveCompanyId, updatedBatch);
        addAuditLog('UPDATE', 'BATCH', prevBatch.id, `Lote editado manualmente: ${prevBatch.id.slice(0, 8)} (Fecha: ${format(new Date(resolvedEntryDate), 'dd/MM/yyyy')}, Stock restante: ${newRemaining})`, prevBatch, updatedBatch);
    }

    setShowModal(null);
    setSelectedBatchId(null);
    setFormData({ 
      quantity: 0, 
      cost: 0, 
      supplier: '', 
      name: '', 
      description: '', 
      category: '', 
      unit: 'litro',
      type: 'materia prima',
      productDescription: ''
    });
  };

  // --- GESTIÓN DE COMPRAS: EDITAR Y ELIMINAR ---
  const handleOpenEditPurchase = (buy: Batch) => {
    setPurchaseToEdit(buy);
    setEditPurchaseForm({
      quantity: buy.quantity,
      unitCost: buy.unitCost,
      supplier: buy.supplier || '',
      entryDate: buy.entryDate.split('T')[0],
      description: buy.description || ''
    });
  };

  const handleSaveEditPurchase = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!purchaseToEdit || !effectiveCompanyId) return;

    const prevBatch = purchaseToEdit;
    const newQty = Number(editPurchaseForm.quantity);
    const newCost = Number(editPurchaseForm.unitCost);
    if (newQty <= 0) {
      alert('La cantidad de compra debe ser mayor a cero.');
      return;
    }
    if (newCost < 0) {
      alert('El costo unitario no puede ser negativo.');
      return;
    }

    const consumed = Math.max(0, prevBatch.quantity - prevBatch.remainingQuantity);
    if (newQty < consumed) {
      alert(`No puede reducir la cantidad comprada a ${newQty} porque ya se han vendido o consumido ${consumed.toFixed(2)} unidades de este lote.`);
      return;
    }

    const newRemaining = Math.max(0, newQty - consumed);
    const resolvedDate = editPurchaseForm.entryDate.includes('T') 
      ? editPurchaseForm.entryDate 
      : `${editPurchaseForm.entryDate}T${prevBatch.entryDate.includes('T') ? prevBatch.entryDate.split('T')[1] : '12:00:00'}`;

    const updatedBatch: Batch = {
      ...prevBatch,
      entryDate: resolvedDate,
      quantity: newQty,
      remainingQuantity: newRemaining,
      unitCost: newCost,
      supplier: editPurchaseForm.supplier.trim() || 'Proveedor',
      description: editPurchaseForm.description.trim()
    };

    await dbService.saveBatch(effectiveCompanyId, updatedBatch);

    const prod = products.find(p => p.id === prevBatch.productId);
    const prodName = prod?.name || 'Producto';

    addAuditLog(
      'UPDATE',
      'BATCH',
      prevBatch.id,
      `Compra editada: Lote ${prevBatch.id.slice(0, 8)} (${prodName}) | Cantidad: ${prevBatch.quantity} -> ${newQty} | Costo: Bs. ${newCost} | Prov: ${updatedBatch.supplier}`,
      prevBatch,
      updatedBatch
    );

    addNotification({
      type: 'SUCCESS',
      title: 'Compra Actualizada',
      message: `Se ha modificado la compra de "${prodName}".`,
      linkTab: 'purchases'
    });

    setPurchaseToEdit(null);
  };

  const handleDeletePurchase = (buy: Batch) => {
    setPurchaseToDelete(buy);
  };

  const executeConfirmDeletePurchase = async () => {
    if (!purchaseToDelete || !effectiveCompanyId || isDeletingMovement) return;
    setIsDeletingMovement(true);

    const buy = purchaseToDelete;
    const prod = products.find(p => p.id === buy.productId);
    const prodName = prod?.name || 'Producto';

    try {
      await dbService.deleteItem(effectiveCompanyId, 'batches', buy.id);

      // Clean from selected purchases if it was marked
      setSelectedPurchases(prev => prev.filter(id => id !== buy.id));

      addAuditLog(
        'DELETE',
        'BATCH',
        buy.id,
        `Compra eliminada: Lote ${buy.id.slice(0, 8)} (${buy.quantity} ${prod?.unit || 'unidad'}s de "${prodName}", Proveedor: ${buy.supplier})`,
        buy,
        null
      );

      addNotification({
        type: 'WARNING',
        title: 'Compra Eliminada',
        message: `Se eliminó el registro de compra de ${buy.quantity} ${prod?.unit || 'unidad'}s de "${prodName}".`,
        linkTab: 'purchases'
      });
      setPurchaseToDelete(null);
    } catch (err) {
      console.error('Error al eliminar compra:', err);
      addNotification({
        type: 'WARNING',
        title: 'Error al Eliminar',
        message: 'No se pudo eliminar el lote de compra. Inténtelo nuevamente.',
        linkTab: 'purchases'
      });
    } finally {
      setIsDeletingMovement(false);
    }
  };

  const executeConfirmDeleteBulkPurchases = async () => {
    if (!bulkPurchasesToDelete || !effectiveCompanyId || isDeletingMovement) return;
    setIsDeletingMovement(true);

    try {
      const items = [...bulkPurchasesToDelete];
      for (const buy of items) {
        await dbService.deleteItem(effectiveCompanyId, 'batches', buy.id);
        const prod = products.find(p => p.id === buy.productId);
        addAuditLog(
          'DELETE',
          'BATCH',
          buy.id,
          `Compra eliminada en lote: ${buy.quantity} ${prod?.unit || 'unidad'}s de "${prod?.name || 'Producto'}"`,
          buy,
          null
        );
      }

      const deletedIds = new Set(items.map(i => i.id));
      setSelectedPurchases(prev => prev.filter(id => !deletedIds.has(id)));

      addNotification({
        type: 'WARNING',
        title: 'Compras Eliminadas',
        message: `Se eliminaron con éxito ${items.length} registros de compras seleccionados.`,
        linkTab: 'purchases'
      });
      setBulkPurchasesToDelete(null);
    } catch (err) {
      console.error('Error al eliminar compras en lote:', err);
      addNotification({
        type: 'WARNING',
        title: 'Error en Eliminación',
        message: 'Ocurrió un error al eliminar algunos registros.',
        linkTab: 'purchases'
      });
    } finally {
      setIsDeletingMovement(false);
    }
  };

  // --- GESTIÓN DE VENTAS: EDITAR, ELIMINAR Y EXPORTAR ---
  const handleOpenEditSale = (sale: Transaction) => {
    setSaleToEdit(sale);
    setEditSaleForm({
      quantity: sale.quantity,
      salePrice: sale.salePrice || sale.unitCost || 0,
      date: sale.date.split('T')[0],
      reference: sale.reference || '',
      description: sale.description || ''
    });
  };

  const handleSaveEditSale = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!saleToEdit || !effectiveCompanyId) return;

    const prevSale = saleToEdit;
    const newQty = Number(editSaleForm.quantity);
    const newSalePrice = Number(editSaleForm.salePrice);
    if (newQty <= 0) {
      alert('La cantidad vendida debe ser mayor a cero.');
      return;
    }
    if (newSalePrice < 0) {
      alert('El precio de venta no puede ser negativo.');
      return;
    }

    const qtyDiff = newQty - prevSale.quantity;

    // If quantity changed, adjust the associated batch
    if (qtyDiff !== 0 && prevSale.batchId) {
      const associatedBatch = batches.find(b => b.id === prevSale.batchId);
      if (associatedBatch) {
        if (qtyDiff > 0 && associatedBatch.remainingQuantity < qtyDiff) {
          alert(`No hay suficiente stock disponible en el lote original para aumentar la venta en ${qtyDiff} unidades (Disponible en lote: ${associatedBatch.remainingQuantity}).`);
          return;
        }
        const updatedBatch = {
          ...associatedBatch,
          remainingQuantity: Math.max(0, Math.min(associatedBatch.quantity, associatedBatch.remainingQuantity - qtyDiff))
        };
        await dbService.saveBatch(effectiveCompanyId, updatedBatch);
      }
    }

    const resolvedDate = editSaleForm.date.includes('T')
      ? editSaleForm.date
      : `${editSaleForm.date}T${prevSale.date.includes('T') ? prevSale.date.split('T')[1] : '12:00:00'}`;

    const updatedTx: Transaction = {
      ...prevSale,
      date: resolvedDate,
      quantity: newQty,
      salePrice: newSalePrice,
      reference: editSaleForm.reference.trim() || prevSale.reference,
      description: editSaleForm.description.trim()
    };

    await dbService.saveTransaction(effectiveCompanyId, updatedTx);

    const prod = products.find(p => p.id === prevSale.productId);
    const prodName = prod?.name || 'Producto';

    addAuditLog(
      'UPDATE',
      'TRANSACTION',
      prevSale.id,
      `Venta editada: ${prodName} | Cantidad: ${prevSale.quantity} -> ${newQty} | Precio: Bs. ${newSalePrice} | Ref: ${updatedTx.reference}`,
      prevSale,
      updatedTx
    );

    addNotification({
      type: 'SUCCESS',
      title: 'Venta Actualizada',
      message: `Se modificó la venta de "${prodName}" (Total: Bs. ${(newQty * newSalePrice).toLocaleString('es-CO', { minimumFractionDigits: 2 })}).`,
      linkTab: 'sales'
    });

    setSaleToEdit(null);
  };

  const handleDeleteSale = (sale: Transaction) => {
    setSaleToDelete(sale);
  };

  const executeConfirmDeleteSale = async () => {
    if (!saleToDelete || !effectiveCompanyId || isDeletingMovement) return;
    setIsDeletingMovement(true);

    const sale = saleToDelete;
    const prod = products.find(p => p.id === sale.productId);
    const prodName = prod?.name || 'Producto';

    try {
      // 1. Restore stock to batch if batchId exists
      let stockRestored = false;
      if (sale.batchId) {
        const associatedBatch = batches.find(b => b.id === sale.batchId);
        if (associatedBatch) {
          const restoredRemaining = Math.min(associatedBatch.quantity, associatedBatch.remainingQuantity + sale.quantity);
          const updatedBatch = {
            ...associatedBatch,
            remainingQuantity: restoredRemaining
          };
          await dbService.saveBatch(effectiveCompanyId, updatedBatch);
          stockRestored = true;
        }
      }

      // If associated batch was not found, restore to another existing batch of the same product or create a return batch
      if (!stockRestored && sale.productId) {
        const candidateBatch = batches.find(b => b.productId === sale.productId && b.remainingQuantity < b.quantity);
        if (candidateBatch) {
          const restoredRemaining = Math.min(candidateBatch.quantity, candidateBatch.remainingQuantity + sale.quantity);
          await dbService.saveBatch(effectiveCompanyId, {
            ...candidateBatch,
            remainingQuantity: restoredRemaining
          });
          stockRestored = true;
        } else {
          // If all batches are full or none found, create a replenishment batch to return stock
          const returnBatch: Batch = {
            id: crypto.randomUUID(),
            productId: sale.productId,
            entryDate: new Date().toISOString(),
            quantity: sale.quantity,
            remainingQuantity: sale.quantity,
            unitCost: sale.unitCost || 0,
            supplier: 'Devolución de Venta',
            description: `Stock restituido por anulación de venta ${sale.id.slice(0, 8)}`,
            type: 'ENTRY',
            performedBy: currentUser?.name || currentUser?.email || 'Sistema'
          };
          await dbService.saveBatch(effectiveCompanyId, returnBatch);
          stockRestored = true;
        }
      }

      // 2. Delete the sale transaction
      await dbService.deleteItem(effectiveCompanyId, 'transactions', sale.id);

      addAuditLog(
        'DELETE',
        'TRANSACTION',
        sale.id,
        `Venta eliminada: ${sale.quantity} ${prod?.unit || 'unidad'}s de "${prodName}" (Stock restituido a almacén)`,
        sale,
        null
      );

      addNotification({
        type: 'SUCCESS',
        title: 'Venta Eliminada',
        message: `Se eliminó la venta de "${prodName}" y se restituyeron ${sale.quantity} ${prod?.unit || 'unidad'}s al inventario.`,
        linkTab: 'sales'
      });

      setSaleToDelete(null);
    } catch (err) {
      console.error('Error al eliminar venta:', err);
      addNotification({
        type: 'WARNING',
        title: 'Error al Eliminar',
        message: 'No se pudo eliminar el registro de venta. Inténtelo nuevamente.',
        linkTab: 'sales'
      });
    } finally {
      setIsDeletingMovement(false);
    }
  };

  const handleExportSalesXLSX = () => {
    if (filteredSaleReports.length === 0) {
      alert('No hay ventas registradas para exportar con los filtros seleccionados.');
      return;
    }
    const headers = ['Fecha', 'Producto', 'Categoría', 'Cantidad', 'Unidad', 'Precio Venta Unit. (Bs)', 'Costo Unit. PEPS (Bs)', 'Total Venta (Bs)', 'Costo Total (Bs)', 'Utilidad Bruta (Bs)', 'Margen (%)', 'Referencia / Factura', 'Notas', 'Vendedor'];
    const rows = filteredSaleReports.map(s => {
      const product = products.find(p => p.id === s.productId);
      const price = s.salePrice || s.unitCost || 0;
      const totalSale = s.quantity * price;
      const totalCost = s.quantity * s.unitCost;
      const profit = totalSale - totalCost;
      const margin = totalSale > 0 ? (profit / totalSale) * 100 : 0;
      return [
        format(new Date(s.date), 'dd/MM/yyyy HH:mm'),
        product?.name || 'Inexistente',
        product?.category || 'N/A',
        s.quantity,
        product?.unit || 'unidad',
        price.toFixed(2),
        s.unitCost.toFixed(2),
        totalSale.toFixed(2),
        totalCost.toFixed(2),
        profit.toFixed(2),
        `${margin.toFixed(1)}%`,
        s.reference || 'N/A',
        s.description || '',
        s.performedBy || 'Sistema'
      ];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Reporte de Ventas');
    XLSX.writeFile(workbook, `reporte_ventas_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
  };

  const handleExportXLSX = (mode: 'filtered' | 'all' = 'filtered') => {
    const finishedProductsList = products.filter(p => p.type === 'producto terminado');
    const entriesToExport = mode === 'all' 
      ? generateUnifiedKardex(finishedProductsList, batches, transactions) 
      : filteredKardex;

    if (entriesToExport.length === 0 && finishedProductsList.length === 0) {
      alert('No hay movimientos ni productos terminados registrados para exportar.');
      return;
    }

    try {
      exportKardexWorkbook(
        entriesToExport,
        finishedProductsList,
        batches,
        {
          companyName: currentUser?.companyName || 'DESTILERÍA ORMÓN',
          userName: currentUser?.name || currentUser?.email || 'Sistema',
          filters: mode === 'filtered' ? {
            product: kardexFilterProduct,
            type: kardexFilterType,
            startDate: kardexFilterStartDate,
            endDate: kardexFilterEndDate,
            search: kardexSearch
          } : undefined,
          includeSummarySheet: true,
          includeBatchesSheet: true,
          fileNamePrefix: mode === 'all' ? 'kardex_peps_productos_terminados' : 'kardex_peps_reporte'
        }
      );

      addNotification({
        type: 'SUCCESS',
        title: 'Kardex Exportado a Excel',
        message: `Se ha generado el libro Excel profesional de Tarjeta Kardex con ${entriesToExport.length} movimientos valorados al Costo Promedio al momento de salida.`,
        linkTab: 'kardex'
      });
    } catch (err) {
      console.error('Error exporting Kardex Excel:', err);
      alert('Ocurrió un error al exportar el archivo Excel: ' + (err instanceof Error ? err.message : String(err)));
    }
  };

  const handleExportSingleProductKardex = (productId: string) => {
    const prod = products.find(p => p.id === productId);
    if (!prod) return;

    const productEntries = generateUnifiedKardex(products, batches, transactions)
      .filter(entry => entry.productId === productId);

    if (productEntries.length === 0) {
      alert(`No existen movimientos registrados en el Kardex para ${prod.name}.`);
      return;
    }

    try {
      exportKardexWorkbook(
        productEntries,
        products,
        batches,
        {
          companyName: currentUser?.companyName || 'DESTILERÍA ORMÓN',
          userName: currentUser?.name || currentUser?.email || 'Sistema',
          specificProduct: prod,
          includeSummarySheet: true,
          includeBatchesSheet: true,
          fileNamePrefix: `kardex_promedio_${prod.name.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()}`
        }
      );

      addNotification({
        type: 'SUCCESS',
        title: 'Kardex de Producto Exportado',
        message: `Se ha generado el archivo Excel de Tarjeta Kardex para ${prod.name} valorado al Costo Promedio al momento de salida.`,
        linkTab: 'kardex'
      });
    } catch (err) {
      console.error('Error exporting single product Kardex:', err);
      alert('Ocurrió un error al exportar el Kardex del producto: ' + (err instanceof Error ? err.message : String(err)));
    }
  };

  const confirmDeleteProduct = () => {
    if (productToDelete && effectiveCompanyId) {
      addAuditLog('DELETE', 'PRODUCT', productToDelete.id, `Producto eliminado: ${productToDelete.name}`, productToDelete, null);
      dbService.deleteItem(effectiveCompanyId, 'products', productToDelete.id);
      
      // Also delete associated batches and transactions for this product
      batches.filter(b => b.productId === productToDelete.id).forEach(b => dbService.deleteItem(effectiveCompanyId, 'batches', b.id));
      transactions.filter(t => t.productId === productToDelete.id).forEach(t => dbService.deleteItem(effectiveCompanyId, 'transactions', t.id));
      
      setProductToDelete(null);
    }
  };

  const handleOpenEditProduct = (prod: Product) => {
    setProductToEdit(prod);
    setEditProductForm({
      name: prod.name,
      description: prod.description || '',
      category: prod.category || '',
      type: prod.type || 'materia prima',
      unit: prod.unit || 'litro',
      minStock: prod.minStock ?? 0,
      recipe: prod.recipe ? [...prod.recipe] : [],
      packaging: prod.packaging ? { ...prod.packaging } : { packagingType: 'ninguno', unitsPerBox: 0, unitsPerBundle: 0 }
    });
  };

  const handleSaveEditProduct = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!productToEdit || !effectiveCompanyId) return;

    const previousProduct = { ...productToEdit };
    const hasRecipe = (editProductForm.type === 'producto terminado' || editProductForm.type === 'producción') && editProductForm.recipe.length > 0;
    const hasPackaging = (editProductForm.type === 'producto terminado' || editProductForm.type === 'producción') && editProductForm.packaging.packagingType !== 'ninguno';

    const updatedProduct: Product = {
      ...productToEdit,
      name: editProductForm.name.trim(),
      description: editProductForm.description.trim(),
      category: editProductForm.category.trim(),
      type: editProductForm.type,
      unit: editProductForm.unit,
      minStock: Number(editProductForm.minStock) || 0,
      recipe: hasRecipe ? editProductForm.recipe : undefined,
      packaging: hasPackaging ? editProductForm.packaging : undefined
    };

    await dbService.saveProduct(effectiveCompanyId, updatedProduct);

    // Immediate UI feedback
    setProducts(prev => prev.map(p => p.id === updatedProduct.id ? updatedProduct : p));

    addAuditLog(
      'UPDATE',
      'PRODUCT',
      updatedProduct.id,
      `Producto editado: "${updatedProduct.name}" | Tipo: ${updatedProduct.type} | Unidad: ${updatedProduct.unit} | Stock Mínimo: ${updatedProduct.minStock}${hasPackaging ? ` | Empaque: ${editProductForm.packaging.packagingType}` : ''}`,
      previousProduct,
      updatedProduct
    );

    addNotification({
      type: 'SUCCESS',
      title: 'Producto Editado',
      message: `El producto "${updatedProduct.name}" ha sido actualizado en la base de datos.`,
      linkTab: 'inventory',
      productId: updatedProduct.id
    });

    setProductToEdit(null);
  };

  const handleExecuteProductionOrder = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!effectiveCompanyId) return;

    const finishedProduct = products.find(p => p.id === produceForm.productId);
    if (!finishedProduct) {
      alert('Seleccione un producto terminado válido.');
      return;
    }

    const recipe = finishedProduct.recipe || [];
    if (recipe.length === 0) {
      alert('Este producto no tiene configurada una receta / consumo de materiales. Edite el producto para definir qué materiales consume.');
      return;
    }

    // Determine total units to produce
    let totalUnits = Number(produceForm.quantityToProduce) || 0;
    if (produceForm.producedIn === 'cajas' && finishedProduct.packaging?.unitsPerBox) {
      totalUnits = (Number(produceForm.boxCount) || 0) * finishedProduct.packaging.unitsPerBox;
    } else if (produceForm.producedIn === 'bultos' && finishedProduct.packaging?.unitsPerBundle) {
      totalUnits = (Number(produceForm.bundleCount) || 0) * finishedProduct.packaging.unitsPerBundle;
    }

    if (totalUnits <= 0) {
      alert('La cantidad a producir debe ser mayor a cero.');
      return;
    }

    // Check material stock requirements
    const missingMaterials: string[] = [];
    for (const item of recipe) {
      const requiredQty = item.quantityNeeded * totalUnits;
      const currentAvailable = productStock[item.materialProductId] || 0;
      if (currentAvailable < requiredQty) {
        const matName = products.find(p => p.id === item.materialProductId)?.name || 'Material desconocido';
        missingMaterials.push(`• ${matName}: se necesitan ${requiredQty.toFixed(2)}, pero solo hay ${currentAvailable.toFixed(2)} en stock.`);
      }
    }

    if (missingMaterials.length > 0) {
      alert(`No hay suficiente materia prima en inventario para llevar a cabo esta producción:\n\n${missingMaterials.join('\n')}\n\nPor favor registre una compra o entrada de estos materiales primero.`);
      return;
    }

    // Process PEPS consumption for each ingredient
    let currentBatches = [...batches];
    let totalProductionCost = 0;
    const consumedDetails: Array<{
      materialProductId: string;
      materialName: string;
      quantityUsed: number;
      unit: string;
      cost: number;
    }> = [];
    const productionTransactions: Transaction[] = [];

    const prodDate = produceForm.productionDate || new Date().toISOString().split('T')[0];
    const companyId = effectiveCompanyId;

    for (const item of recipe) {
      const requiredQty = item.quantityNeeded * totalUnits;
      const mat = products.find(p => p.id === item.materialProductId);
      const matName = mat?.name || 'Materia Prima';

      const { updatedBatches, transactions: txs, totalCost: ingredientCost } = processExit(
        requiredQty,
        currentBatches,
        item.materialProductId,
        `Consumo para producir ${totalUnits} ${finishedProduct.unit}s de "${finishedProduct.name}"`,
        'EXIT',
        currentUser.name || currentUser.email || 'Producción',
        undefined,
        prodDate
      );

      currentBatches = updatedBatches;
      totalProductionCost += ingredientCost;
      productionTransactions.push(...txs);

      consumedDetails.push({
        materialProductId: item.materialProductId,
        materialName: matName,
        quantityUsed: requiredQty,
        unit: mat?.unit || item.unit || 'unidad',
        cost: ingredientCost
      });
    }

    // Calculate unit cost for the newly created batch
    const unitCost = totalUnits > 0 ? (totalProductionCost / totalUnits) : 0;

    // Create the finished product batch in inventory
    const boxesCount = finishedProduct.packaging?.unitsPerBox && finishedProduct.packaging.unitsPerBox > 0 
      ? Math.floor(totalUnits / finishedProduct.packaging.unitsPerBox) 
      : undefined;
    const bundlesCount = finishedProduct.packaging?.unitsPerBundle && finishedProduct.packaging.unitsPerBundle > 0 
      ? Math.floor(totalUnits / finishedProduct.packaging.unitsPerBundle) 
      : undefined;

    let packagingDesc = '';
    if (produceForm.producedIn === 'cajas') {
      packagingDesc = ` (${produceForm.boxCount} cajas de ${finishedProduct.packaging?.unitsPerBox} ${finishedProduct.unit}s c/u)`;
    } else if (produceForm.producedIn === 'bultos') {
      packagingDesc = ` (${produceForm.bundleCount} bultos de ${finishedProduct.packaging?.unitsPerBundle} ${finishedProduct.unit}s c/u)`;
    } else if (boxesCount || bundlesCount) {
      packagingDesc = ` (Equivale a ${boxesCount ? `${boxesCount} cajas` : ''}${boxesCount && bundlesCount ? ' / ' : ''}${bundlesCount ? `${bundlesCount} bultos` : ''})`;
    }

    const nowTime = new Date().toTimeString().split(' ')[0];
    const resolvedProdDate = prodDate.includes('T') ? prodDate : `${prodDate}T${nowTime}`;

    const newFinishedBatch: Batch = {
      id: crypto.randomUUID(),
      productId: finishedProduct.id,
      entryDate: resolvedProdDate,
      quantity: totalUnits,
      remainingQuantity: totalUnits,
      unitCost: unitCost,
      supplier: 'Producción Interna',
      description: `Lote fabricado a partir de materias primas${packagingDesc}. ${produceForm.notes || ''}`.trim(),
      type: 'ENTRY',
      performedBy: currentUser?.name || currentUser?.email || 'Producción'
    };

    // Save production order record
    const orderRecord: ProductionOrderRecord = {
      id: crypto.randomUUID(),
      finishedProductId: finishedProduct.id,
      finishedProductName: finishedProduct.name,
      quantityProduced: totalUnits,
      boxesProduced: boxesCount,
      bundlesProduced: bundlesCount,
      totalCost: totalProductionCost,
      unitCost: unitCost,
      date: new Date(prodDate + 'T12:00:00').toISOString(),
      performedBy: currentUser?.name || currentUser?.email || 'Producción',
      consumedMaterials: consumedDetails,
      notes: produceForm.notes
    };

    // Commit changes to Firebase
    // 1. Updated ingredient batches
    currentBatches
      .filter(b => {
        const oldB = batches.find(ob => ob.id === b.id);
        return oldB && oldB.remainingQuantity !== b.remainingQuantity;
      })
      .forEach(b => dbService.saveBatch(companyId, b));

    // 2. Ingredient consumption transactions
    productionTransactions.forEach(tx => dbService.saveTransaction(companyId, tx));

    // 3. New finished product batch
    await dbService.saveBatch(companyId, newFinishedBatch);

    // 4. Production order record
    await dbService.saveProductionOrder(companyId, orderRecord);

    // 5. Audit logs
    addAuditLog(
      'CREATE',
      'PRODUCTION',
      orderRecord.id,
      `Producción completada: ${totalUnits} ${finishedProduct.unit}s de "${finishedProduct.name}" (Costo total: $${totalProductionCost.toLocaleString('es-CO', { minimumFractionDigits: 2 })}). Materiales consumidos: ${consumedDetails.map(d => `${d.quantityUsed} ${d.unit} de ${d.materialName}`).join(', ')}`,
      null,
      orderRecord
    );

    addNotification({
      type: 'SUCCESS',
      title: '¡Producción Exitosa!',
      message: `Se fabricaron ${totalUnits} ${finishedProduct.unit}s de "${finishedProduct.name}" ingresadas al inventario. Materias primas descontadas por PEPS.`,
      linkTab: 'inventory',
      productId: finishedProduct.id
    });

    setShowModal(null);
    setProduceForm({
      productId: '',
      quantityToProduce: 1,
      producedIn: 'unidades',
      boxCount: 0,
      bundleCount: 0,
      notes: '',
      productionDate: new Date().toISOString().split('T')[0]
    });
  };

  const handlePeopleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPerson.name || !newPerson.email || !effectiveCompanyId) return;

    if (newPerson.id) {
      // Edit mode
      const prev = team.find(t => t.id === newPerson.id);
      const updated = { ...prev, ...newPerson, email: newPerson.email?.trim().toLowerCase() } as TeamMember;
      dbService.saveTeamMember(effectiveCompanyId, updated);
      addAuditLog('UPDATE', 'TEAM', newPerson.id, `Colaborador actualizado: ${newPerson.name}`, prev, updated);
    } else {
      // Create mode
      const person: TeamMember = {
        id: crypto.randomUUID(),
        name: newPerson.name!,
        email: newPerson.email!.trim().toLowerCase(),
        role: (newPerson.role || 'OPERATOR') as UserRole,
        areas: (newPerson.areas || []) as AreaPermission[],
        status: (newPerson.status || 'active') as 'active' | 'inactive'
      };
      dbService.saveTeamMember(effectiveCompanyId, person);
      addAuditLog('CREATE', 'TEAM', person.id, `Nuevo colaborador añadido: ${person.name}`, null, person);
    }

    setShowPeopleModal(false);
    setNewPerson({ areas: [], role: 'OPERATOR', status: 'active' });
  };

  const handleConsolidatePurchases = () => {
    if (selectedPurchases.length < 2) return;
    
    const selectedBatches = batches.filter(b => selectedPurchases.includes(b.id));
    const firstProduct = selectedBatches[0].productId;
    const sameProduct = selectedBatches.every(b => b.productId === firstProduct);
    
    if (!sameProduct) {
      alert("Solo se pueden consolidar compras del mismo producto.");
      return;
    }

    // Sort by date to find the most recent
    const sortedByDate = [...selectedBatches].sort((a, b) => 
      new Date(b.entryDate).getTime() - new Date(a.entryDate).getTime()
    );

    // Check if they are all from the same calendar day
    const allSameDay = selectedBatches.every(b => 
      format(new Date(b.entryDate), 'yyyy-MM-dd') === format(new Date(selectedBatches[0].entryDate), 'yyyy-MM-dd')
    );

    // Use specific date if all same day, otherwise use the most recent
    const entryDate = allSameDay ? selectedBatches[0].entryDate : sortedByDate[0].entryDate;

    // Determine weighted average cost based on remaining quantity (or total quantity if exhausted)
    const totalQty = selectedBatches.reduce((sum, b) => sum + b.quantity, 0);
    const totalRemaining = selectedBatches.reduce((sum, b) => sum + b.remainingQuantity, 0);
    const weightedCost = totalRemaining > 0 
      ? selectedBatches.reduce((sum, b) => sum + (b.remainingQuantity * b.unitCost), 0) / totalRemaining
      : (totalQty > 0 ? selectedBatches.reduce((sum, b) => sum + (b.quantity * b.unitCost), 0) / totalQty : 0);
    
    const suppliers = Array.from(new Set(selectedBatches.map(b => b.supplier))).join(', ');

    const consolidatedBatch: Batch = {
      id: crypto.randomUUID(),
      productId: firstProduct,
      entryDate: entryDate,
      quantity: totalQty,
      remainingQuantity: totalRemaining,
      unitCost: weightedCost,
      supplier: suppliers.length > 30 ? suppliers.slice(0, 27) + '...' : suppliers, 
      description: `Consolidación: ${selectedBatches.length} entradas${allSameDay ? ' del mismo día' : ''}`,
      type: 'PURCHASE',
      performedBy: currentUser?.name || currentUser?.email || 'Sistema'
    };

    if (effectiveCompanyId) {
      selectedPurchases.forEach(id => dbService.deleteItem(effectiveCompanyId, 'batches', id));
      dbService.saveBatch(effectiveCompanyId, consolidatedBatch);
    }

    addAuditLog('SYSTEM', 'BATCH', consolidatedBatch.id, `Compras consolidadas (${selectedBatches.length} registros) para el producto ${products.find(p => p.id === firstProduct)?.name}`);
    addNotification({
      type: 'SUCCESS',
      title: 'Compras Consolidadas',
      message: `Se han consolidado ${selectedBatches.length} compras. Fecha: ${format(new Date(entryDate), 'dd/MM/yyyy')}`,
      linkTab: 'purchases'
    });

    setSelectedPurchases([]);
  };

  const [toasts, setToasts] = useState<AppNotification[]>([]);

  const addToast = (notif: Omit<AppNotification, 'id' | 'timestamp' | 'read'>) => {
    const id = Math.random().toString(36).substr(2, 9);
    const toast: AppNotification = {
      ...notif,
      id,
      timestamp: new Date().toISOString(),
      read: false
    };
    setToasts(prev => [...prev, toast]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 5000);
  };

  const handleEditPerson = (member: TeamMember) => {
    setNewPerson(member);
    setShowPeopleModal(true);
  };

  const toggleArea = (area: AreaPermission) => {
    const current = newPerson.areas || [];
    if (current.includes(area)) {
      setNewPerson({ ...newPerson, areas: current.filter(a => a !== area) });
    } else {
      setNewPerson({ ...newPerson, areas: [...current, area] });
    }
  };

  // Components
  const NotificationCenter = () => (
    <AnimatePresence>
      {showNotifications && (
        <>
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setShowNotifications(false)}
            className="fixed inset-0 bg-black/20 backdrop-blur-[2px] z-[90]"
          />
          <motion.div 
            initial={{ opacity: 0, x: 20, scale: 0.95 }}
            animate={{ opacity: 1, x: 0, scale: 1 }}
            exit={{ opacity: 0, x: 20, scale: 0.95 }}
            className="fixed top-20 right-6 lg:top-24 lg:right-12 w-full max-w-sm bg-card border border-main rounded-3xl shadow-2xl z-[100] overflow-hidden flex flex-col max-h-[600px]"
          >
            <div className="p-6 border-b border-main flex items-center justify-between bg-sidebar/50">
              <div>
                <h3 className="text-sm font-black uppercase tracking-widest text-main">Centro de Notificaciones</h3>
                <p className="text-[9px] font-black uppercase tracking-widest text-muted opacity-60">Control de alertas y sistema</p>
              </div>
              <button 
                onClick={clearNotifications}
                className="text-[9px] font-black uppercase tracking-widest text-red-500 hover:bg-red-500/10 px-3 py-1.5 rounded-lg transition-colors border border-red-500/20"
              >
                Limpiar Todo
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-4 space-y-3 custom-scrollbar">
              {notifications.length === 0 ? (
                <div className="py-20 text-center">
                  <Bell size={32} className="text-muted/20 mx-auto mb-3" />
                  <p className="text-[10px] font-black uppercase tracking-widest text-muted opacity-40">No hay notificaciones pendientes</p>
                </div>
              ) : (
                notifications.map(n => (
                  <div 
                    key={n.id} 
                    className={cn(
                      "p-4 rounded-2xl border transition-all relative group",
                      n.read ? "bg-sidebar/30 border-main/10 opacity-70" : "bg-sidebar border-accent/20 shadow-sm"
                    )}
                    onClick={() => {
                      markAsRead(n.id);
                      if (n.linkTab) setActiveTab(n.linkTab as any);
                    }}
                  >
                    {!n.read && (
                      <span className="absolute top-4 right-4 w-2 h-2 bg-accent rounded-full" />
                    )}
                    <div className="flex gap-3">
                      <div className={cn(
                        "w-10 h-10 rounded-xl flex items-center justify-center shrink-0",
                        n.type === 'LOW_STOCK' ? "bg-red-500/10 text-red-500" : "bg-green-500/10 text-green-500"
                      )}>
                        {n.type === 'LOW_STOCK' ? <AlertTriangle size={18} /> : <CheckCircle size={18} />}
                      </div>
                      <div className="flex-1">
                        <h4 className="text-[11px] font-black uppercase tracking-wider text-main mb-1">{n.title}</h4>
                        <p className="text-[10px] font-medium text-muted leading-relaxed line-clamp-2">{n.message}</p>
                        <span className="text-[8px] font-bold text-muted/40 uppercase tracking-widest mt-2 block">
                          {format(new Date(n.timestamp), 'HH:mm • dd MMM', { locale: es })}
                        </span>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>

            {notifications.length > 0 && notifications.some(n => !n.read) && (
              <div className="p-4 border-t border-main bg-sidebar/30">
                <button 
                  onClick={() => setNotifications(prev => prev.map(n => ({ ...n, read: true })))}
                  className="w-full bg-accent/10 text-accent py-3 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-accent hover:text-white transition-all border border-accent/20"
                >
                  Marcar todas como leídas
                </button>
              </div>
            )}
          </motion.div>
        </>
      )}
    </AnimatePresence>
  );

  const SidebarItem = ({ id, icon: Icon, label }: { id: typeof activeTab, icon: any, label: string }) => (
    <button
      onClick={() => {
        setActiveTab(id);
        setIsSidebarOpen(false);
      }}
      className={cn(
        "flex items-center gap-4 w-full px-4 py-4 rounded-xl transition-all duration-300 group",
        activeTab === id 
          ? "bg-accent text-white shadow-xl shadow-brand-gold-shadow font-black" 
          : "text-muted hover:text-white hover:bg-white/5"
      )}
    >
      <Icon size={22} className={cn(activeTab === id ? "text-white" : "text-muted group-hover:text-white")} />
      <span className="text-sm tracking-tight">{label}</span>
    </button>
  );

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-sidebar flex items-center justify-center p-6 selection:bg-accent selection:text-white font-inter">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-md bg-white rounded-[2.5rem] p-12 shadow-2xl shadow-brand-gold-shadow/10 border border-main/5 text-center"
        >
          <div className="flex flex-col items-center mb-10">
             <div className="w-24 h-24 mb-6">
                <img 
                    src="/logo-sied.jpeg" 
                    alt="Logo SIED" 
                   className="w-full h-full object-contain"
                />
             </div>
             <h2 className="text-2xl font-black text-main uppercase tracking-[0.1em] mb-1">
               {isForgotPassword ? 'Cambiar Contraseña' :
                registerForm === 'empresa' ? 'Registro Empresa' : 
                registerForm === 'usuario' ? 'Registro Usuario' : 
                'Iniciar Sesión'}
             </h2>
             <span className="text-[10px] font-bold text-muted uppercase tracking-[0.3em]">
               {isForgotPassword ? 'Seguridad y Acceso' : 'Sied System'}
             </span>
          </div>

          <AnimatePresence mode="wait">
            {isForgotPassword ? (
              <motion.form 
                key="forgot-password-form"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                onSubmit={handleResetPassword}
                className="space-y-6 text-left"
              >
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                    Correo electrónico
                  </label>
                  <div className="relative group">
                    <Mail size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    <input 
                      type="email" 
                      required
                      value={resetEmail}
                      onChange={(e) => {
                        setResetEmail(e.target.value);
                        setResetError('');
                      }}
                      placeholder="tu@correo.com"
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-4 outline-none transition-all font-bold text-sm"
                    />
                  </div>
                  <p className="text-[9px] font-bold text-muted/60 uppercase tracking-wider ml-1">Correo de la cuenta a modificar</p>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                    Contraseña nueva
                  </label>
                  <div className="relative group">
                    <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    <input 
                      type={showNewPassword ? "text" : "password"} 
                      required
                      value={newPassword}
                      onChange={(e) => {
                        setNewPassword(e.target.value);
                        setResetError('');
                      }}
                      placeholder="••••••••"
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-12 outline-none transition-all font-bold text-sm"
                    />
                    <button
                      type="button"
                      onClick={() => setShowNewPassword(!showNewPassword)}
                      className="absolute right-4 top-1/2 -translate-y-1/2 text-muted hover:text-main transition-colors p-1"
                      title={showNewPassword ? "Ocultar" : "Mostrar"}
                    >
                      {showNewPassword ? <EyeOff size={18} /> : <Eye size={18} />}
                    </button>
                  </div>
                  <p className="text-[9px] font-bold text-muted/60 uppercase tracking-wider ml-1">Mínimo 6 caracteres</p>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                    Confirmar contraseña
                  </label>
                  <div className="relative group">
                    <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    <input 
                      type={showConfirmPassword ? "text" : "password"} 
                      required
                      value={confirmPassword}
                      onChange={(e) => {
                        setConfirmPassword(e.target.value);
                        setResetError('');
                      }}
                      placeholder="••••••••"
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-12 outline-none transition-all font-bold text-sm"
                    />
                    <button
                      type="button"
                      onClick={() => setShowConfirmPassword(!showConfirmPassword)}
                      className="absolute right-4 top-1/2 -translate-y-1/2 text-muted hover:text-main transition-colors p-1"
                      title={showConfirmPassword ? "Ocultar" : "Mostrar"}
                    >
                      {showConfirmPassword ? <EyeOff size={18} /> : <Eye size={18} />}
                    </button>
                  </div>
                </div>

                {resetError && (
                  <div className="p-4 bg-red-500/10 border border-red-500/20 text-red-500 text-xs font-bold rounded-2xl flex items-center gap-2.5">
                    <AlertTriangle size={18} className="shrink-0" />
                    <span>{resetError}</span>
                  </div>
                )}

                {resetSuccess && (
                  <div className="p-4 bg-emerald-500/10 border border-emerald-500/20 text-emerald-600 text-xs font-bold rounded-2xl flex items-center gap-2.5">
                    <CheckCircle size={18} className="shrink-0 text-emerald-500" />
                    <span>¡Contraseña actualizada! La clave anterior fue borrada de la base de datos. Ingresando únicamente con la nueva clave...</span>
                  </div>
                )}

                <button 
                  type="submit"
                  disabled={isResettingPassword || resetSuccess}
                  className={cn(
                    "w-full bg-accent hover:bg-accent/90 text-white font-black py-5 rounded-[1.5rem] shadow-xl shadow-brand-gold-shadow transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-xs mt-4 flex items-center justify-center gap-2",
                    (isResettingPassword || resetSuccess) && "opacity-50 cursor-not-allowed"
                  )}
                >
                  {isResettingPassword ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      Actualizando...
                    </>
                  ) : 'Cambiar Contraseña'}
                </button>

                <button 
                  type="button"
                  onClick={() => {
                    setIsForgotPassword(false);
                    setResetError('');
                    setResetSuccess(false);
                    setResetEmail('');
                    setNewPassword('');
                    setConfirmPassword('');
                  }}
                  className="w-full text-[10px] font-black uppercase tracking-widest text-muted hover:text-main transition-colors text-center block pt-2"
                >
                  Volver al inicio de sesión
                </button>
              </motion.form>
            ) : registerForm ? (
              <motion.form 
                key="register-form"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                onSubmit={handleRegister}
                className="space-y-6 text-left"
              >
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                    {registerForm === 'empresa' ? 'Nombre de la Empresa' : 'Nombre Completo'}
                  </label>
                  <div className="relative group">
                    {registerForm === 'empresa' ? (
                      <Building2 size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    ) : (
                      <User size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    )}
                    <input 
                      type="text" 
                      required
                      value={registerForm === 'empresa' ? regFormData.company : regFormData.name}
                      onChange={(e) => setRegFormData({ ...regFormData, [registerForm === 'empresa' ? 'company' : 'name']: e.target.value })}
                      placeholder={registerForm === 'empresa' ? "Nombre de la empresa" : "ej. Juan Pérez"}
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-6 outline-none transition-all font-bold text-sm"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Correo Electrónico</label>
                  <div className="relative group">
                    <Mail size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    <input 
                      type="email" 
                      required
                      value={regFormData.email}
                      onChange={(e) => setRegFormData({ ...regFormData, email: e.target.value })}
                      placeholder="correo@ejemplo.com"
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-6 outline-none transition-all font-bold text-sm"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Contraseña</label>
                  <div className="relative group">
                    <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                    <input 
                      type="password" 
                      required
                      value={regFormData.password}
                      onChange={(e) => setRegFormData({ ...regFormData, password: e.target.value })}
                      placeholder="••••••••"
                      className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-6 outline-none transition-all font-bold text-sm"
                    />
                  </div>
                </div>

                {registerForm === 'usuario' && (
                  <div className="bg-accent/5 border border-accent/20 rounded-2xl p-4 text-xs font-bold text-accent/90 leading-relaxed flex items-start gap-2.5">
                    <span className="text-sm mt-0.5">ℹ️</span>
                    <span>
                      Si fuiste agregado previamente como personal de una empresa, regístrate con ese mismo correo electrónico para vincularte automáticamente.
                    </span>
                  </div>
                )}

                <button 
                  type="submit"
                  disabled={isRegistering}
                  className={cn(
                    "w-full bg-accent hover:bg-accent/90 text-white font-black py-5 rounded-[1.5rem] shadow-xl shadow-brand-gold-shadow transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-xs mt-4 flex items-center justify-center gap-2",
                    isRegistering && "opacity-50 cursor-not-allowed"
                  )}
                >
                  {isRegistering ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      Registrando...
                    </>
                  ) : 'Confirmar Registro'}
                </button>

                <div className="flex items-center gap-4 my-2">
                  <div className="flex-1 h-px bg-zinc-100" />
                  <span className="text-zinc-400 text-[10px] font-bold uppercase tracking-widest">o</span>
                  <div className="flex-1 h-px bg-zinc-100" />
                </div>

                <button 
                  type="button"
                  onClick={async () => {
                    const type = registerForm === 'empresa' ? 'COMPANY' : 'USER';
                    const companyName = registerForm === 'empresa' ? regFormData.company : undefined;
                    
                    if (type === 'COMPANY' && !companyName) {
                      alert('Por favor, ingresa el nombre de tu empresa en el primer campo antes de continuar con Google.');
                      return;
                    }

                    try {
                      setIsRegistering(true);
                      const profile = await authService.loginWithGoogle(type, companyName);
                      if (profile) setCurrentUser(profile);
                    } catch (err: any) {
                      console.error('Google registration error:', err);
                      const errorCode = err.code || '';
                      const errorMessage = err.message || '';
                      
                      if (errorCode === 'auth/operation-not-allowed' || errorMessage.includes('auth/operation-not-allowed')) {
                        alert('ERROR: El método de inicio de sesión con Google está DESACTIVADO en Firebase.\n\nActívalo en la Consola de Firebase: https://console.firebase.google.com/project/' + (firebaseConfig as any).projectId + '/authentication/providers');
                      } else {
                        alert('Error al registrar con Google.');
                      }
                    } finally {
                      setIsRegistering(false);
                    }
                  }}
                  className="w-full bg-white border border-zinc-200 text-main font-bold py-4 rounded-[1.5rem] shadow-sm transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-[10px] flex items-center justify-center gap-3 hover:bg-zinc-50"
                >
                  <svg className="w-4 h-4" viewBox="0 0 24 24">
                    <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                    <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-1 .67-2.28 1.07-3.71 1.07-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                    <path fill="#FBBC05" d="M5.84 14.11c-.22-.66-.35-1.36-.35-2.11s.13-1.45.35-2.11V7.05H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.95l3.66-2.84z" />
                    <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.05l3.66 2.84c.87-2.6 3.3-4.51 6.16-4.51z" />
                  </svg>
                  Continuar con Google
                </button>

                <button 
                  type="button"
                  onClick={() => setRegisterForm(null)}
                  className="w-full text-[10px] font-black uppercase tracking-widest text-muted hover:text-main transition-colors text-center"
                >
                  Cancelar y volver
                </button>
              </motion.form>
            ) : (
              <motion.div 
                key="login-or-options"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 20 }}
              >
                {!showRegisterOptions ? (
                  <form 
                    onSubmit={handleLogin}
                    className="space-y-6 text-left"
                  >
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Usuario / Correo</label>
                      <div className="relative group">
                        <User size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                        <input 
                          type="email" 
                          required
                          value={loginFormData.email}
                          onChange={(e) => setLoginFormData({ ...loginFormData, email: e.target.value })}
                          placeholder="correo@ejemplo.com"
                          className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-6 outline-none transition-all font-bold text-sm"
                        />
                      </div>
                    </div>

                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Contraseña</label>
                      <div className="relative group">
                        <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-accent opacity-50 group-focus-within:opacity-100 transition-opacity" />
                        <input 
                          type="password" 
                          required
                          value={loginFormData.password}
                          onChange={(e) => setLoginFormData({ ...loginFormData, password: e.target.value })}
                          placeholder="••••••••"
                          className="w-full bg-sidebar border border-transparent focus:border-accent/40 rounded-2xl py-4 pl-12 pr-6 outline-none transition-all font-bold text-sm"
                        />
                      </div>
                    </div>

                    <div className="flex items-center justify-between py-2">
                      <label className="flex items-center gap-2 cursor-pointer group">
                        <input type="checkbox" className="w-4 h-4 rounded-md border-main accent-accent" />
                        <span className="text-xs font-bold text-muted group-hover:text-main transition-colors">Recordarme</span>
                      </label>
                      <button 
                        type="button" 
                        onClick={() => {
                          setIsForgotPassword(true);
                          setResetEmail(loginFormData.email.trim());
                          setResetError('');
                          setResetSuccess(false);
                          setNewPassword('');
                          setConfirmPassword('');
                        }}
                        className="text-[10px] font-black uppercase tracking-widest text-accent hover:underline decoration-2 underline-offset-4"
                      >
                        ¿Olvidaste tu contraseña?
                      </button>
                    </div>

                <button 
                  type="submit"
                  disabled={isLoggingIn}
                  className={cn(
                    "w-full bg-accent hover:bg-accent/90 text-white font-black py-5 rounded-[1.5rem] shadow-xl shadow-brand-gold-shadow transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-xs mt-4 flex items-center justify-center gap-2",
                    isLoggingIn && "opacity-50 cursor-not-allowed"
                  )}
                >
                  {isLoggingIn ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      Accediendo...
                    </>
                  ) : 'Acceder al Sistema'}
                </button>

                <div className="flex items-center gap-4 my-4">
                  <div className="flex-1 h-px bg-zinc-800" />
                  <span className="text-zinc-500 text-[10px] font-bold uppercase tracking-widest">o</span>
                  <div className="flex-1 h-px bg-zinc-800" />
                </div>

                <button 
                  type="button"
                  onClick={async () => {
                    try {
                      setIsLoggingIn(true);
                      const profile = await authService.loginWithGoogle();
                      if (profile) setCurrentUser(profile);
                    } catch (err: any) {
                      console.error('Google login error detail:', err);
                      const errorCode = err.code || '';
                      const errorMessage = err.message || '';

                      if (errorCode === 'auth/operation-not-allowed' || errorMessage.includes('auth/operation-not-allowed')) {
                        alert('ERROR: El método de inicio de sesión con Google está DESACTIVADO en Firebase.\n\nActívalo en la Consola de Firebase: https://console.firebase.google.com/project/' + (firebaseConfig as any).projectId + '/authentication/providers');
                      } else {
                        alert('Error al acceder con Google. Asegúrese de que las ventanas emergentes estén permitidas.');
                      }
                    } finally {
                      setIsLoggingIn(false);
                    }
                  }}
                  className="w-full bg-white text-black font-black py-4 rounded-[1.5rem] shadow-lg transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-[10px] flex items-center justify-center gap-3"
                >
                  <svg className="w-4 h-4" viewBox="0 0 24 24">
                    <path fill="currentColor" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                    <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-1 .67-2.28 1.07-3.71 1.07-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                    <path fill="#FBBC05" d="M5.84 14.11c-.22-.66-.35-1.36-.35-2.11s.13-1.45.35-2.11V7.05H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.95l3.66-2.84z" />
                    <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.05l3.66 2.84c.87-2.6 3.3-4.51 6.16-4.51z" />
                  </svg>
                  Continuar con Google
                </button>

                    <div className="relative pt-4 text-center">
                      <button 
                        type="button"
                        onClick={() => setShowRegisterOptions(true)}
                        className="w-full bg-white border-2 border-accent text-accent font-black py-4 rounded-[1.5rem] transition-all active:scale-[0.98] uppercase tracking-[0.2em] text-[10px] hover:bg-accent hover:text-white"
                      >
                        Crear Cuenta Nueva
                      </button>
                    </div>
                  </form>
                ) : (
                  <div className="space-y-6">
                    <div className="grid grid-cols-2 gap-3">
                      <button 
                        type="button"
                        onClick={() => setShowRegisterOptions(false)}
                        className="col-span-2 text-[9px] font-black uppercase tracking-widest text-muted mb-2 flex items-center justify-center gap-1 hover:text-main"
                      >
                        <ChevronRight size={12} className="rotate-180" /> Volver al Login
                      </button>
                      <button 
                        type="button"
                        onClick={() => setRegisterForm('empresa')}
                        className="flex flex-col items-center justify-center gap-3 p-4 bg-sidebar border border-main/10 rounded-3xl hover:border-accent transition-all group"
                      >
                        <div className="w-12 h-12 rounded-2xl bg-accent/10 flex items-center justify-center text-accent group-hover:bg-accent group-hover:text-white transition-all transform group-active:scale-95">
                          <Warehouse size={22} />
                        </div>
                        <span className="text-[10px] font-black uppercase tracking-widest">Empresa</span>
                      </button>
                      <button 
                        type="button"
                        onClick={() => setRegisterForm('usuario')}
                        className="flex flex-col items-center justify-center gap-3 p-4 bg-sidebar border border-main/10 rounded-3xl hover:border-accent transition-all group"
                      >
                        <div className="w-12 h-12 rounded-2xl bg-accent/10 flex items-center justify-center text-accent group-hover:bg-accent group-hover:text-white transition-all transform group-active:scale-95">
                          <UserCircle size={22} />
                        </div>
                        <span className="text-[10px] font-black uppercase tracking-widest">Usuario</span>
                      </button>
                    </div>
                  </div>
                )}
              </motion.div>
            )}
          </AnimatePresence>

          <p className="text-center text-[9px] font-bold text-muted uppercase tracking-widest mt-10 opacity-40">
            © 2026 SIED SYSTEM • Todos los derechos reservados
          </p>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-main text-main font-inter selection:bg-accent selection:text-white overflow-hidden">
      {/* Mobile Top Bar */}
      <div className="lg:hidden fixed top-0 left-0 right-0 h-16 bg-sidebar border-b border-main flex items-center justify-between px-4 z-50">
        <div className="flex items-center gap-3">
          <Menu 
            className="text-accent cursor-pointer active:scale-95 transition-transform" 
            onClick={() => setIsSidebarOpen(true)}
            size={24}
          />
          <div className="flex flex-col">
            <h1 className="font-black text-xs tracking-wider text-accent uppercase leading-none">{companyName || 'SIED'}</h1>
            {availableCompanies.length > 1 && (
              <select
                value={effectiveCompanyId}
                onChange={(e) => {
                  const selected = availableCompanies.find(c => c.id === e.target.value);
                  if (selected) {
                    handleSwitchCompany(selected.id, selected.name);
                  }
                }}
                className="bg-transparent text-[10px] font-bold text-muted outline-none cursor-pointer mt-0.5"
              >
                {availableCompanies.map(c => (
                  <option key={c.id} value={c.id} className="bg-neutral-900 text-white">
                    {c.name} ({c.productCount} prod.)
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div className="text-[10px] font-bold text-emerald-500 bg-emerald-500/10 px-2 py-0.5 rounded border border-emerald-500/20">
            {products.length} prod.
          </div>
          <button 
            onClick={() => setShowNotifications(!showNotifications)}
            className="relative p-2 text-muted hover:text-accent transition-colors"
          >
            <Bell size={20} />
            {notifications.filter(n => !n.read).length > 0 && (
              <span className="absolute top-1.5 right-1.5 w-2 h-2 bg-red-500 rounded-full border-2 border-sidebar" />
            )}
          </button>
          <div className="w-8 h-8 rounded-full bg-accent/10 flex items-center justify-center text-accent">
            <UserCircle size={20} />
          </div>
        </div>
      </div>

      {/* Sidebar Overlay (Mobile) */}
      <AnimatePresence>
        {isSidebarOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setIsSidebarOpen(false)}
            className="lg:hidden fixed inset-0 bg-black/60 backdrop-blur-sm z-[60]"
          />
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside className={cn(
        "fixed inset-y-0 left-0 w-80 bg-sidebar border-r border-main flex flex-col p-8 z-[70] transition-transform duration-300 transform lg:translate-x-0 lg:static lg:inset-auto",
        isSidebarOpen ? "translate-x-0 shadow-2xl" : "-translate-x-full"
      )}>
        <div className="flex items-center justify-between mb-14 lg:block lg:mb-14">
          <div className="flex items-center gap-5 px-1">
            <div className="w-16 h-16 lg:w-20 lg:h-20 flex items-center justify-center">
              <img 
                 src="/logo-sied.png" 
                 alt="SIED Logo" 
                 className="w-full h-full object-contain"
                 onError={(e) => {
                   e.currentTarget.style.display = 'none';
                   e.currentTarget.parentElement!.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="w-12 h-12 text-accent"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/></svg>`;
                 }}
              />
            </div>
            <div className="flex flex-col justify-center overflow-hidden">
              <h1 className="font-black text-sm lg:text-[17px] tracking-[0.05em] text-accent uppercase leading-none truncate">
                {companyName ? companyName.split(' ')[0] : 'SIED'}
              </h1>
              <h1 className="font-black text-sm lg:text-[17px] tracking-[0.05em] text-accent uppercase leading-tight truncate">
                {companyName ? companyName.split(' ').slice(1).join(' ') : 'SYSTEM'}
              </h1>
              <span className="text-[8px] lg:text-[9px] font-black text-muted tracking-widest uppercase opacity-60 mt-2 block">SIED SYSTEM</span>
            </div>
          </div>
          <button 
            onClick={() => setIsSidebarOpen(false)}
            className="lg:hidden p-2 text-muted hover:text-accent transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        <nav className="flex-1 space-y-2 lg:space-y-3 overflow-y-auto pr-2 custom-scrollbar">
          <SidebarItem id="dashboard" icon={LayoutDashboard} label="Panel Control" />
          
          {accountType === 'COMPANY' && (
            <>
              <SidebarItem id="inventory" icon={Package} label="Inventario" />
              <SidebarItem id="purchases" icon={ShoppingCart} label="Compras" />
              <SidebarItem id="sales" icon={BadgeDollarSign} label="Ventas" />
              <SidebarItem id="kardex" icon={ClipboardList} label="Tarjeta Kardex" />
              <SidebarItem id="audit" icon={Activity} label="Trazabilidad" />
              <SidebarItem id="people" icon={UserCircle} label="Personal" />
              <SidebarItem id="configuration" icon={Settings} label="Configuración" />
            </>
          )}

          {accountType === 'USER' && (
              <>
                  {userAreas.includes('INVENTARIO') && <SidebarItem id="inventory" icon={Package} label="Inventario" />}
                  <SidebarItem id="purchases" icon={ShoppingCart} label="Compras" />
                  <SidebarItem id="sales" icon={BadgeDollarSign} label="Ventas" />
                  {userAreas.includes('REPORTES') && <SidebarItem id="kardex" icon={ClipboardList} label="Tarjeta Kardex" />}
                  {role === 'ADMIN' && <SidebarItem id="audit" icon={Activity} label="Trazabilidad" />}
                  {userAreas.includes('CONFIGURACION') && (
                    <>
                      <SidebarItem id="people" icon={UserCircle} label="Personal" />
                      <SidebarItem id="configuration" icon={Settings} label="Configuración" />
                    </>
                  )}
              </>
          )}
        </nav>

        <div className="pt-8 border-t border-main mt-4">
          <button 
            onClick={handleLogout}
            className="flex items-center gap-4 w-full px-4 py-3 text-red-500 hover:bg-red-500/10 rounded-xl transition-all group"
          >
            <LogOut size={22} className="group-hover:-translate-x-1 transition-transform" />
            <span className="font-black text-sm uppercase tracking-widest">Cerrar Sesión</span>
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto pt-24 lg:pt-0 transition-all flex flex-col h-full">
        {/* Desktop Top Bar */}
        <div className="hidden lg:flex sticky top-0 bg-main/80 backdrop-blur-md px-12 py-5 items-center justify-between z-40 border-b border-main/5">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <h2 className="text-xs font-black uppercase text-muted tracking-[0.2em]">Dashboard</h2>
              <ChevronRight size={12} className="text-muted/30" />
              <span className="text-xs font-black uppercase text-accent tracking-[0.2em]">{activeTab}</span>
            </div>

            {/* Company / Workspace Switcher */}
            {availableCompanies.length > 0 && (
              <div className="flex items-center gap-2 bg-sidebar border border-main px-3 py-1.5 rounded-xl ml-4">
                <Building2 size={16} className="text-accent" />
                <div className="flex flex-col">
                  <span className="text-[8px] font-black uppercase tracking-widest text-muted">Empresa / Base de Datos</span>
                  <select
                    value={effectiveCompanyId}
                    onChange={(e) => {
                      const selected = availableCompanies.find(c => c.id === e.target.value);
                      if (selected) {
                        handleSwitchCompany(selected.id, selected.name);
                      }
                    }}
                    className="bg-transparent text-xs font-black uppercase text-accent outline-none cursor-pointer tracking-wider"
                  >
                    {availableCompanies.map(c => (
                      <option key={c.id} value={c.id} className="bg-neutral-900 text-white font-bold">
                        {c.name} ({c.productCount} {c.productCount === 1 ? 'producto' : 'productos'}, {c.batchCount} lotes)
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}

            <div className="hidden xl:flex items-center gap-2 px-3 py-1 rounded-lg bg-emerald-500/10 border border-emerald-500/20 text-emerald-500 text-[10px] font-bold">
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
              <span>Firebase Conectado • {products.length} productos • {batches.length} lotes</span>
            </div>
          </div>
          <div className="flex items-center gap-6">
            <button 
              onClick={() => setShowNotifications(!showNotifications)}
              className="relative p-3 bg-sidebar border border-main rounded-2xl text-muted hover:text-accent hover:border-accent transition-all group"
            >
              <Bell size={20} className="group-hover:rotate-12 transition-transform" />
              {notifications.filter(n => !n.read).length > 0 && (
                <span className="absolute -top-1 -right-1 min-w-[20px] h-5 bg-red-500 text-white text-[10px] font-black rounded-full flex items-center justify-center border-4 border-main animate-bounce">
                  {notifications.filter(n => !n.read).length}
                </span>
              )}
            </button>
            <div className="flex items-center gap-4 pl-6 border-l border-main/10">
              <div className="text-right">
                <p className="text-[10px] font-black uppercase tracking-widest text-main">{currentUser?.name || currentUser?.email}</p>
                <p className="text-[8px] font-black uppercase tracking-widest text-muted">{role}</p>
              </div>
              <div className="w-10 h-10 rounded-2xl bg-accent/10 flex items-center justify-center text-accent border border-accent/20">
                <User size={20} />
              </div>
            </div>
          </div>
        </div>

        <div className="p-12 space-y-8 flex-1 overflow-y-auto custom-scrollbar">
          <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8"
            >
              <header>
                <h2 className="text-3xl font-extrabold tracking-tight mb-4">
                  Bienvenido, {accountType === 'COMPANY' ? 'Administrador de Empresa' : 'Colaborador'}
                </h2>
                <div className="flex gap-4 mt-2 text-sm text-muted">
                    <span>Sistema: <span className="text-accent font-semibold tracking-wide uppercase">{companyName || 'Sied System'}</span></span>
                    <span className="opacity-20 text-main/20">|</span>
                    <span>Modo: <span className="text-accent font-semibold">PEPS Activo</span></span>
                </div>
              </header>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-sidebar p-6 rounded-2xl border border-main shadow-sm">
                  <div className="text-xs font-bold uppercase tracking-widest text-muted mb-2">Stock Total</div>
                  <div className="text-3xl font-bold">{(Object.values(productStock) as number[]).reduce((a, b) => a + b, 0)}</div>
                </div>
                <div className="bg-sidebar p-6 rounded-2xl border border-main shadow-sm">
                  <div className="text-xs font-bold uppercase tracking-widest text-muted mb-2">Líneas de Producto</div>
                  <div className="text-3xl font-bold">{products.length}</div>
                </div>
                {products.some(p => productStock[p.id] <= p.minStock) && (
                  <div className="bg-red-500/10 p-6 rounded-2xl border border-red-500/30 shadow-sm cursor-pointer" onClick={() => setActiveTab('inventory')}>
                    <div className="flex items-center gap-2 mb-2">
                      <AlertTriangle size={14} className={cn(
                        "text-red-500",
                        products.some(p => productStock[p.id] <= p.minStock * criticalThreshold) && "animate-pulse"
                      )} />
                      <div className="text-xs font-black uppercase tracking-widest text-red-500 whitespace-nowrap">Alertas de Stock</div>
                    </div>
                    <div className="text-3xl font-extrabold text-red-600">
                      {products.filter(p => productStock[p.id] <= p.minStock).length} <span className="text-sm font-bold opacity-60">ALERTA</span>
                    </div>
                  </div>
                )}
                {accountType === 'COMPANY' && (
                  <>
                    <div className="bg-sidebar p-6 rounded-2xl border border-main shadow-sm">
                      <div className="text-xs font-bold uppercase tracking-widest text-muted mb-2">Valoración Inventario (Bs.)</div>
                      <div className="text-3xl font-bold text-accent">Bs. {dashboardStats.valuation.toLocaleString()}</div>
                    </div>
                    <div className="bg-sidebar p-6 rounded-2xl border border-main shadow-sm border-l-red-500 border-l-4">
                      <div className="text-xs font-bold uppercase tracking-widest text-muted mb-2">Costo de Ventas Histórico</div>
                      <div className="text-3xl font-bold text-red-500">Bs. {dashboardStats.cogs.toLocaleString()}</div>
                    </div>
                  </>
                )}
              </div>

              <div className="bg-sidebar rounded-2xl border border-main overflow-hidden shadow-sm">
                <div className="px-6 py-4 border-b border-main flex justify-between items-center bg-black/5 dark:bg-white/5">
                  <h3 className="font-bold uppercase tracking-widest text-xs text-muted">Estado del Existencias</h3>
                  <button 
                    onClick={() => {
                      const headers = ['Producto', 'Categoría', 'Tipo', 'Stock Actual', 'Unidad', 'Mínimo'];
                      const rows = products.map(p => [
                        p.name,
                        p.category,
                        p.type,
                        productStock[p.id],
                        p.unit,
                        p.minStock
                      ]);
                      const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                      const workbook = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(workbook, worksheet, 'Inventario');
                      XLSX.writeFile(workbook, `inventario_actual_${format(new Date(), 'yyyyMMdd')}.xlsx`);
                    }}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-main text-muted hover:text-main hover:bg-main/5 transition-all font-black uppercase text-[9px] tracking-widest"
                  >
                    <Download size={12} />
                    Exportar XLSX
                  </button>
                </div>
                <table className="w-full text-left">
                  <thead className="text-muted text-[10px] font-bold uppercase tracking-wider">
                    <tr>
                      <th className="px-6 py-3 border-b border-main/20">Producto</th>
                      <th className="px-6 py-3 border-b border-main/20">Categoría</th>
                      <th className="px-6 py-3 border-b border-main/20">Existencia</th>
                      <th className="px-6 py-3 border-b border-main/20">Unidad</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-main/5">
                    {products.map(product => (
                      <tr key={product.id} className="hover:bg-black/5 dark:hover:bg-white/5 transition-colors">
                        <td className="px-6 py-4 font-semibold text-main">
                          <div className="flex items-center gap-2">
                            {product.name}
                            {productStock[product.id] <= product.minStock && (
                              <div className="flex items-center gap-1.5">
                                {productStock[product.id] <= product.minStock * criticalThreshold ? (
                                  <AlertTriangle size={14} className="text-red-500 animate-pulse" title="STOCK CRÍTICO" />
                                ) : (
                                  <div className="w-2 h-2 rounded-full bg-red-500" title="Bajo el umbral mínimo" />
                                )}
                              </div>
                            )}
                          </div>
                        </td>
                        <td className="px-6 py-4 text-muted">{product.category}</td>
                        <td className="px-6 py-4">
                          <span className={cn(
                            "px-3 py-1 rounded font-bold text-xs uppercase tracking-tighter",
                            productStock[product.id] > 0 
                              ? "bg-green-500/10 text-green-500 border border-green-500/20" 
                              : "bg-red-500/10 text-red-500 border border-red-500/20"
                          )}>
                            {productStock[product.id]}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-muted">{product.unit}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )}

          {activeTab === 'people' && (
            <motion.div 
              key="people"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8"
            >
               <header className="flex flex-col lg:flex-row justify-between items-start lg:items-center bg-sidebar p-8 rounded-[2.5rem] border border-main shadow-sm gap-8 transition-all">
                <div className="flex-1">
                  <div className="flex items-center gap-3 mb-2">
                    <Users className="text-accent" size={24} />
                    <h2 className="text-3xl font-black uppercase tracking-tighter">Gestión de Personal</h2>
                  </div>
                  <p className="text-muted text-sm uppercase tracking-widest font-bold opacity-60">Control de accesos, roles y permisos por áreas operativas</p>
                </div>
                
                <div className="flex flex-wrap items-center gap-6">
                   <div className="flex flex-col items-end px-6 border-r border-main/20">
                      <span className="text-[10px] font-black uppercase text-muted/60 tracking-widest">Total Equipo</span>
                      <span className="text-2xl font-black text-accent">{team.length}</span>
                   </div>
                   <div className="flex flex-col items-end px-6 border-r border-main/20">
                      <span className="text-[10px] font-black uppercase text-muted/60 tracking-widest">Admins</span>
                      <span className="text-2xl font-black text-main">{team.filter(t => t.role === 'ADMIN').length}</span>
                   </div>
                   <div className="flex flex-col items-end px-6">
                      <span className="text-[10px] font-black uppercase text-muted/60 tracking-widest">Operadores</span>
                      <span className="text-2xl font-black text-main">{team.filter(t => t.role === 'OPERATOR').length}</span>
                   </div>
                   <button 
                    onClick={() => { setNewPerson({ areas: [], role: 'OPERATOR', status: 'active' }); setShowPeopleModal(true); }}
                    className="bg-accent text-white px-10 py-5 rounded-[1.5rem] font-black flex items-center gap-3 hover:brightness-110 active:scale-95 transition-all shadow-2xl shadow-brand-gold-shadow text-xs uppercase tracking-widest ml-4"
                  >
                    <Plus size={20} /> Añadir Colaborador
                  </button>
                </div>
              </header>

              {/* Filters Bar */}
              <div className="flex flex-col md:flex-row gap-4">
                <div className="relative flex-1">
                  <Search className="absolute left-6 top-1/2 -translate-y-1/2 text-muted" size={18} />
                  <input 
                    type="text"
                    placeholder="Buscar por nombre o correo..."
                    className="w-full bg-sidebar border border-main rounded-2xl pl-16 pr-6 py-5 focus:outline-none focus:border-accent font-bold text-main placeholder:text-muted/40 transition-all uppercase text-xs tracking-widest"
                    value={peopleSearch}
                    onChange={e => setPeopleSearch(e.target.value)}
                  />
                </div>
                <div className="flex gap-4">
                  <select 
                    className="bg-sidebar border border-main rounded-2xl px-6 py-5 focus:outline-none focus:border-accent font-bold text-main appearance-none uppercase text-xs tracking-widest min-w-[180px]"
                    value={peopleRoleFilter}
                    onChange={e => setPeopleRoleFilter(e.target.value as any)}
                  >
                    <option value="ALL">TODOS LOS ROLES</option>
                    <option value="ADMIN">ADMINISTRADORES</option>
                    <option value="OPERATOR">OPERADORES</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {filteredTeam.map(member => (
                   <div key={member.id} className={cn(
                     "bg-white p-8 rounded-[2.5rem] border border-main flex flex-col gap-6 group hover:border-accent/40 transition-all shadow-sm hover:shadow-xl relative overflow-hidden",
                     member.status === 'inactive' && "opacity-60 grayscale"
                   )}>
                      <div className="flex justify-between items-start">
                        <div className="flex items-center gap-5">
                           <div className={cn(
                             "w-16 h-16 rounded-2xl flex items-center justify-center transition-colors",
                             member.role === 'ADMIN' ? "bg-accent/10 text-accent" : "bg-main/5 text-muted"
                           )}>
                              {member.role === 'ADMIN' ? <Shield size={28} /> : <User size={28} />}
                           </div>
                           <div>
                              <div className="flex items-center gap-3">
                                <h3 className="text-xl font-black uppercase tracking-tight">{member.name}</h3>
                                {member.status === 'inactive' && (
                                  <span className="text-[8px] font-black uppercase bg-red-500 text-white px-2 py-0.5 rounded-full tracking-[0.2em]">Inactivo</span>
                                )}
                              </div>
                              <p className="text-muted text-sm font-bold tracking-tight opacity-70 mb-2">{member.email}</p>
                              <div className="flex items-center gap-4">
                               <div className="flex items-center gap-1.5 grayscale opacity-80">
                                  <Shield size={12} className="text-accent" />
                                  <span className="text-[10px] font-black uppercase tracking-[0.2em]">{member.role}</span>
                               </div>
                               <span className="w-1 h-1 rounded-full bg-main/20"></span>
                               <span className="text-[10px] font-black uppercase tracking-[0.2em] text-muted">{member.areas.length} Áreas</span>
                              </div>
                           </div>
                        </div>

                        <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-all translate-y-2 group-hover:translate-y-0">
                           <button 
                            onClick={() => handleEditPerson(member)}
                            className="p-4 hover:bg-main/5 rounded-2xl transition-colors text-muted hover:text-main" title="Editar Colaborador"
                           >
                              <Pencil size={18} />
                           </button>
                           <button 
                            onClick={() => {
                              const updatedStatus = member.status === 'active' ? 'inactive' : 'active';
                              setTeam(team.map(t => t.id === member.id ? { ...t, status: updatedStatus } as TeamMember : t));
                              addAuditLog('UPDATE', 'TEAM', member.id, `Estado cambiado a ${updatedStatus} para: ${member.name}`);
                            }}
                            className={cn(
                              "p-4 rounded-2xl transition-colors",
                              member.status === 'active' ? "text-muted hover:text-orange-500 hover:bg-orange-500/10" : "text-green-500 hover:bg-green-500/10"
                            )}
                            title={member.status === 'active' ? 'Desactivar' : 'Activar'}
                           >
                              {member.status === 'active' ? <X size={18} /> : <CheckCircle size={18} />}
                           </button>
                           <button 
                            onClick={() => {
                              if (confirm(`¿Seguro que desea eliminar a ${member.name}? Se perderá su historial como responsable.`)) {
                                setTeam(team.filter(t => t.id !== member.id));
                                addAuditLog('DELETE', 'TEAM', member.id, `Colaborador eliminado: ${member.name}`);
                              }
                            }}
                            className="p-4 hover:bg-red-500/10 rounded-2xl transition-colors text-muted hover:text-red-500" title="Eliminar del Equipo"
                           >
                              <Trash2 size={18} />
                           </button>
                        </div>
                      </div>

                      <div className="pt-6 border-t border-main/10">
                         <div className="text-[9px] font-black uppercase tracking-widest text-muted mb-4 opacity-60">Permisos por Área Operativa</div>
                         <div className="flex flex-wrap gap-2">
                            {member.areas.map(area => (
                               <span key={area} className="text-[9px] font-bold uppercase tracking-widest bg-sidebar border border-main rounded-xl px-4 py-2 flex items-center gap-2 group/area hover:border-accent/40 transition-all cursor-default">
                                  <div className="w-1.5 h-1.5 rounded-full bg-accent" />
                                  {area}
                               </span>
                            ))}
                            {member.areas.length === 0 && <span className="text-[10px] font-bold italic text-muted opacity-50 uppercase tracking-widest">Sin áreas de acceso definidas</span>}
                         </div>
                      </div>
                   </div>
                ))}

                {filteredTeam.length === 0 && (
                  <div className="lg:col-span-2 py-20 text-center bg-sidebar rounded-[2.5rem] border border-dashed border-main">
                    <Users size={48} className="text-muted/20 mx-auto mb-4" />
                    <p className="text-muted font-bold uppercase tracking-widest text-xs">No se encontraron colaboradores con estos criterios</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'inventory' && (
            <motion.div 
              key="inventory"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-6"
            >
              <header className="flex flex-col md:flex-row justify-between items-start md:items-center bg-sidebar p-6 rounded-2xl border border-main shadow-sm mb-8 gap-6">
                <div>
                  <div className="flex items-center gap-3">
                    <h2 className="text-2xl md:text-3xl font-black uppercase tracking-tighter">Operaciones de Planta</h2>
                    <span className="bg-accent/10 text-accent border border-accent/20 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-widest">
                      {companyName || 'Empresa'}
                    </span>
                  </div>
                  <p className="text-muted text-[10px] md:text-xs uppercase tracking-widest font-bold mt-1">
                    Gestión Central de Almacén • {products.length} productos registrados • {batches.length} lotes en Firebase
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
                  <button 
                    onClick={() => { 
                      setSelectedProductId(null); 
                      setExitItems([]);
                      setFormData({ 
                        quantity: 0, 
                        cost: 0, 
                        minStock: 0, 
                        supplier: '', 
                        name: '', 
                        description: '', 
                        category: '', 
                        unit: 'litro', 
                        type: 'materia prima', 
                        productDescription: 'Salida de almacén', 
                        entryDate: new Date().toISOString().split('T')[0] 
                      }); 
                      setShowModal('exit'); 
                    }}
                    className="flex-1 md:flex-initial bg-sidebar border border-amber-500/40 text-amber-500 hover:bg-amber-500/10 px-5 py-3 rounded-xl font-bold flex items-center justify-center gap-2 transition-all text-[10px] md:text-xs uppercase tracking-widest"
                  >
                    <Minus size={16} /> Registrar Salida / Múltiple
                  </button>
                  <button 
                    onClick={() => {
                      const finishedProds = products.filter(p => p.type === 'producto terminado' || p.type === 'producción');
                      const firstFinished = finishedProds[0];
                      setProduceForm({
                        productId: firstFinished ? firstFinished.id : (products[0]?.id || ''),
                        quantityToProduce: 1,
                        producedIn: 'unidades',
                        boxCount: 1,
                        bundleCount: 1,
                        notes: '',
                        productionDate: new Date().toISOString().split('T')[0]
                      });
                      setShowModal('produce');
                    }}
                    className="flex-1 md:flex-initial bg-emerald-600 hover:bg-emerald-500 text-white px-5 py-3 rounded-xl font-bold flex items-center justify-center gap-2 hover:brightness-110 active:scale-95 transition-all shadow-lg shadow-emerald-950/20 text-[10px] md:text-xs uppercase tracking-widest"
                    title="Transformar materias primas en productos terminados"
                  >
                    <Factory size={16} /> Fabricar / Producir
                  </button>
                  {accountType === 'COMPANY' && (
                    <button 
                      onClick={() => { 
                        setProductRecipe([]);
                        setProductPackaging({ packagingType: 'ninguno', unitsPerBox: 0, unitsPerBundle: 0 });
                        setShowModal('product'); 
                        setFormData({ quantity: 0, cost: 0, minStock: 0, supplier: '', name: '', description: '', category: '', unit: 'unidad', type: 'producto terminado', productDescription: '', entryDate: new Date().toISOString().split('T')[0] }); 
                      }}
                      className="flex-1 md:flex-initial bg-accent text-white px-6 py-3 rounded-xl font-bold flex items-center justify-center gap-2 hover:brightness-110 active:scale-95 transition-all shadow-lg shadow-brand-gold-shadow text-[10px] md:text-xs uppercase tracking-widest"
                    >
                      <Plus size={16} /> Nuevo Producto
                    </button>
                  )}
                  <button 
                    onClick={handleSeedSystemData}
                    disabled={isSeedingData}
                    className="flex-1 md:flex-initial bg-sidebar border border-muted/30 text-muted hover:text-accent hover:border-accent/40 px-4 py-3 rounded-xl font-bold flex items-center justify-center gap-2 transition-all text-[10px] md:text-xs uppercase tracking-widest shadow-sm disabled:opacity-50"
                    title="Cargar productos de demostración de Destilería Ormón (Opcional)"
                  >
                    <Sparkles size={16} className={isSeedingData ? 'animate-spin' : ''} />
                    <span>{isSeedingData ? 'Cargando Demo...' : 'Catálogo Demo'}</span>
                  </button>
                </div>
              </header>

              {/* Search and Filters */}
              <div className="flex flex-col md:flex-row gap-4 mb-8 bg-sidebar p-4 rounded-2xl border border-main shadow-sm">
                <div className="relative flex-1">
                  <input 
                    type="text"
                    placeholder="Buscar producto o categoría..."
                    value={inventorySearch}
                    onChange={(e) => setInventorySearch(e.target.value)}
                    className="w-full bg-main/5 border border-main/20 rounded-xl pl-12 pr-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                  />
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-muted/40" size={16} />
                </div>
                
                {products.some(p => productStock[p.id] <= p.minStock) && (
                  <div className="flex flex-col sm:flex-row gap-2">
                    <button 
                      onClick={() => setShowLowStockOnly(!showLowStockOnly)}
                      className={cn(
                        "flex-1 flex items-center justify-center gap-1.5 px-4 py-3 sm:py-1.5 rounded-xl border transition-all font-black uppercase text-[10px] tracking-widest",
                        showLowStockOnly 
                          ? "bg-red-500 text-white border-red-600" 
                          : "bg-red-500/10 text-red-500 border-red-500/20 hover:bg-red-500/20"
                      )}
                    >
                      <AlertTriangle size={12} />
                      <span>Alertas: {products.filter(p => productStock[p.id] <= p.minStock).length}</span>
                    </button>
                    <button 
                      onClick={() => {
                        const lowStock = products.filter(p => productStock[p.id] <= p.minStock);
                        const headers = ['Producto', 'Categoría', 'Stock Actual', 'Mínimo', 'Estado'];
                        const rows = lowStock.map(p => {
                          const stock = productStock[p.id];
                          const status = stock <= p.minStock * criticalThreshold ? 'CRÍTICO' : 'BAJO';
                          return [p.name, p.category, stock, p.minStock, status];
                        });
                        
                        const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                        const workbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(workbook, worksheet, 'Alertas');
                        XLSX.writeFile(workbook, `alertas_stock_${format(new Date(), 'yyyyMMdd')}.xlsx`);
                      }}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl border border-main bg-sidebar text-muted hover:bg-main/5 transition-all font-black uppercase text-[10px] tracking-widest"
                    >
                      <Download size={12} />
                      <span>Excel Reporte</span>
                    </button>
                  </div>
                )}
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {products
                  .filter(p => !showLowStockOnly || productStock[p.id] <= p.minStock)
                  .filter(p => {
                    const s = inventorySearch.toLowerCase();
                    return !inventorySearch || p.name.toLowerCase().includes(s) || p.category.toLowerCase().includes(s);
                  })
                  .map(product => (
                  <div key={product.id} className="group bg-sidebar p-8 rounded-2xl border border-main shadow-sm hover:border-accent/30 transition-all">
                    <div className="flex justify-between items-start mb-6">
                      <div>
                        <span className="text-[10px] font-bold uppercase tracking-[0.2em] text-accent mb-1 block">{product.category}</span>
                        <div className="flex items-center gap-3">
                          <h3 className="text-2xl font-bold tracking-tight">{product.name}</h3>
                          {productStock[product.id] <= product.minStock && (
                            <div className={cn(
                              "flex items-center gap-1.5 px-2 py-0.5 rounded shadow-sm border",
                              productStock[product.id] <= product.minStock * criticalThreshold 
                                ? "bg-red-600 text-white border-red-700 animate-pulse" 
                                : "bg-red-500/10 text-red-500 border-red-500/20"
                            )}>
                              <AlertTriangle size={12} />
                              <span className="text-[9px] font-black uppercase tracking-widest">
                                {productStock[product.id] <= product.minStock * criticalThreshold ? 'CRÍTICO' : 'AVISO'}
                              </span>
                            </div>
                          )}
                          <button 
                            onClick={() => handleOpenEditProduct(product)}
                            className="p-1.5 hover:bg-accent/20 text-accent rounded-lg transition-colors opacity-70 group-hover:opacity-100"
                            title="Editar Producto (Nombre, Tipo, Unidad...)"
                          >
                            <Pencil size={16} />
                          </button>
                          {accountType === 'COMPANY' && (
                            <button 
                              onClick={() => setProductToDelete(product)}
                              className="p-1.5 hover:bg-red-500/20 text-red-500 rounded-lg transition-colors opacity-40 group-hover:opacity-100"
                              title="Eliminar Producto"
                            >
                              <Trash2 size={16} />
                            </button>
                          )}
                        </div>
                        <p className="text-muted mt-1 text-sm italic opacity-70">{product.description}</p>
                      </div>
                      <div className="text-right">
                        <div className="text-3xl font-black text-main">{productStock[product.id]}</div>
                        <div className="text-[10px] font-black text-muted uppercase tracking-widest">{product.unit}s</div>
                        <div className="text-[8px] font-black text-muted/40 uppercase tracking-tighter mt-1 italic">Mínimo: {product.minStock}</div>
                      </div>
                    </div>
                    <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
                      <div className="flex flex-wrap items-center gap-2">
                        <span className={cn(
                          "text-[9px] font-black uppercase tracking-widest px-2 py-1 rounded border",
                          product.type === 'producto terminado' 
                            ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/30"
                            : product.type === 'producción'
                            ? "bg-purple-500/10 text-purple-500 border-purple-500/30"
                            : "bg-main/5 dark:bg-white/10 text-muted border-main"
                        )}>
                          {product.type}
                        </span>
                        <span className="text-[9px] font-black uppercase tracking-widest bg-accent/10 text-accent px-2 py-1 rounded border border-accent/20 font-mono">
                          {product.unit}s
                        </span>
                        {product.packaging && product.packaging.packagingType !== 'ninguno' && (
                          <span className="text-[9px] font-bold uppercase tracking-wider bg-blue-500/10 text-blue-500 px-2 py-1 rounded border border-blue-500/20 flex items-center gap-1">
                            <Boxes size={10} />
                            {product.packaging.packagingType === 'caja' && `Caja (${product.packaging.unitsPerBox} ${product.unit}s)`}
                            {product.packaging.packagingType === 'bulto' && `Bulto (${product.packaging.unitsPerBundle} ${product.unit}s)`}
                            {product.packaging.packagingType === 'ambos' && `Caja (${product.packaging.unitsPerBox}) / Bulto (${product.packaging.unitsPerBundle})`}
                          </span>
                        )}
                        {product.recipe && product.recipe.length > 0 && (
                          <span className="text-[9px] font-bold uppercase tracking-wider bg-amber-500/10 text-amber-500 px-2 py-1 rounded border border-amber-500/20 flex items-center gap-1" title={`${product.recipe.length} materiales requeridos`}>
                            <Layers size={10} />
                            {product.recipe.length} materias primas
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-2">
                        {(product.type === 'producto terminado' || product.type === 'producción') && (
                          <button 
                            onClick={() => {
                              setProduceForm({
                                productId: product.id,
                                quantityToProduce: 1,
                                producedIn: 'unidades',
                                boxCount: 1,
                                bundleCount: 1,
                                notes: '',
                                productionDate: new Date().toISOString().split('T')[0]
                              });
                              setShowModal('produce');
                            }}
                            className="text-xs font-black text-emerald-500 hover:text-white hover:bg-emerald-600 flex items-center gap-1.5 py-1 px-2.5 rounded-lg bg-emerald-500/10 transition-all border border-emerald-500/30"
                            title="Fabricar este producto terminado descontando materias primas"
                          >
                            <Factory size={12} />
                            <span>Fabricar</span>
                          </button>
                        )}
                        <button 
                          onClick={() => handleOpenEditProduct(product)}
                          className="text-xs font-bold text-accent hover:brightness-125 flex items-center gap-1.5 py-1 px-2.5 rounded-lg bg-accent/10 hover:bg-accent/20 transition-all border border-accent/20"
                          title="Editar nombre, tipo, unidad, receta y empaque..."
                        >
                          <Pencil size={12} />
                          <span>Editar</span>
                        </button>
                      </div>
                    </div>

                    {/* Quick packaging stock breakdown if applicable */}
                    {product.packaging && product.packaging.packagingType !== 'ninguno' && productStock[product.id] > 0 && (
                      <div className="mb-4 p-2.5 rounded-xl bg-blue-500/5 border border-blue-500/15 flex flex-wrap items-center justify-between text-[11px] gap-2">
                        <span className="text-muted font-bold text-[10px] uppercase tracking-wider flex items-center gap-1.5">
                          <PackageCheck size={12} className="text-blue-500" /> Conversión de Empaque:
                        </span>
                        <div className="flex items-center gap-3 font-mono font-bold text-main">
                          {product.packaging.unitsPerBox && product.packaging.unitsPerBox > 0 && (
                            <span className="bg-blue-500/10 px-2 py-0.5 rounded text-blue-600 dark:text-blue-400">
                              📦 {(productStock[product.id] / product.packaging.unitsPerBox).toFixed(1)} Cajas ({product.packaging.unitsPerBox} c/u)
                            </span>
                          )}
                          {product.packaging.unitsPerBundle && product.packaging.unitsPerBundle > 0 && (
                            <span className="bg-indigo-500/10 px-2 py-0.5 rounded text-indigo-600 dark:text-indigo-400">
                              🧺 {(productStock[product.id] / product.packaging.unitsPerBundle).toFixed(1)} Bultos ({product.packaging.unitsPerBundle} c/u)
                            </span>
                          )}
                        </div>
                      </div>
                    )}

                    <div className="flex flex-col sm:flex-row gap-2">
                      <button 
                        onClick={() => { 
                          setSelectedProductId(product.id); 
                          setShowModal('entry'); 
                          setFormData({ 
                            quantity: 0, 
                            cost: 0, 
                            supplier: '', 
                            name: product.name, 
                            description: '', 
                            category: product.category, 
                            unit: product.unit, 
                            type: product.type, 
                            productDescription: 'Entrada directa',
                            entryDate: new Date().toISOString().split('T')[0]
                          }); 
                        }}
                        className="flex-[2] flex items-center justify-center gap-2 bg-accent text-white py-3 md:py-4 rounded-xl font-bold hover:brightness-110 active:scale-95 transition-all shadow-lg shadow-brand-gold-shadow text-sm"
                      >
                        <Plus size={18} /> Entrada
                      </button>
                      <button 
                        onClick={() => { 
                          setSelectedProductId(product.id); 
                          setExitItems([]);
                          setShowModal('exit'); 
                          setFormData({ 
                            quantity: 0, 
                            cost: 0, 
                            supplier: '', 
                            name: product.name, 
                            description: '', 
                            category: product.category, 
                            unit: product.unit, 
                            type: product.type, 
                            productDescription: 'Salida de almacén',
                            entryDate: new Date().toISOString().split('T')[0]
                          }); 
                        }}
                        className="flex-[2] flex items-center justify-center gap-2 bg-transparent border-2 border-accent text-accent py-3 md:py-4 rounded-xl font-bold hover:bg-accent/5 transition-all text-sm"
                      >
                        <Minus size={18} /> Salida
                      </button>
                      {product.type === 'producto terminado' ? (
                        <button 
                          onClick={() => { setSelectedProductId(product.id); setShowModal('product_history'); }}
                          className="flex-[1.5] flex items-center justify-center gap-1.5 px-3 py-3 md:py-4 bg-emerald-500/10 border border-emerald-500/30 text-emerald-600 dark:text-emerald-400 hover:bg-emerald-500 hover:text-white rounded-xl transition-all text-xs font-black uppercase tracking-wider shadow-sm active:scale-95"
                          title="Ver Tarjeta Kardex Valorada de este Producto Terminado"
                        >
                          <ClipboardList size={16} /> Tarjeta Kardex
                        </button>
                      ) : (
                        <button 
                          onClick={() => { setSelectedProductId(product.id); setShowModal('product_history'); }}
                          className="flex-1 flex items-center justify-center p-3 md:p-4 bg-sidebar border border-main text-muted hover:text-main rounded-xl transition-all"
                          title="Ver Historial de Movimientos"
                        >
                          <History size={18} />
                        </button>
                      )}
                    </div>

                    <div className="mt-8 pt-6 border-t border-main/30">
                      <div className="flex justify-between items-center mb-4">
                        <h4 className="text-[10px] font-bold text-muted uppercase tracking-[0.2em]">Prioridad PEPS (Cola)</h4>
                        <span className="bg-accent/10 text-accent text-[9px] font-bold px-2 py-0.5 rounded border border-accent/20 tracking-widest">ACTIVO</span>
                      </div>
                      <div className="space-y-2">
                        {batches.filter(b => b.productId === product.id && b.remainingQuantity > 0)
                          .sort((a, b) => new Date(a.entryDate).getTime() - new Date(b.entryDate).getTime())
                          .map((batch, idx) => (
                            <div key={batch.id} className="flex justify-between items-center text-sm p-4 bg-black/5 dark:bg-white/5 rounded-xl border border-main group/batch">
                              <div className="flex flex-col">
                                <div className="flex items-center gap-2">
                                    <span className={cn(
                                        "text-[9px] font-bold px-1.5 py-0.5 rounded",
                                        idx === 0 ? "bg-accent text-white" : "bg-main/10 text-muted"
                                    )}>{idx === 0 ? 'PRÓX. SALIDA' : `LOTE ${idx + 1}`}</span>
                                    <span className="font-bold text-main">ID: {batch.id.slice(0, 5)}</span>
                                    <button 
                                      onClick={() => {
                                        setSelectedBatchId(batch.id);
                                        setFormData({
                                          ...formData,
                                          quantity: batch.quantity,
                                          cost: batch.unitCost,
                                          supplier: batch.supplier,
                                          productDescription: batch.description || '',
                                          entryDate: batch.entryDate ? (batch.entryDate.includes('T') ? batch.entryDate.split('T')[0] : batch.entryDate) : new Date().toISOString().split('T')[0]
                                        });
                                        setShowModal('edit_batch');
                                      }}
                                      className="p-1.5 hover:bg-accent/20 text-accent rounded-lg transition-all opacity-0 group-hover/batch:opacity-100 flex items-center gap-1.5 border border-accent/20 bg-accent/5 ml-auto"
                                      title="Editar Lote"
                                    >
                                      <Pencil size={11} />
                                      <span className="text-[9px] font-black uppercase tracking-tighter">Editar</span>
                                    </button>
                                </div>
                                <span className="text-[10px] text-muted mt-1 uppercase tracking-wider">Recibido: {format(new Date(batch.entryDate), 'dd MMM yy', { locale: es })}</span>
                                {batch.description && (
                                  <span className="text-[9px] text-accent font-black mt-1 uppercase italic truncate max-w-[150px]">{batch.description}</span>
                                )}
                              </div>
                              <div className="text-right">
                                <div className="font-black text-accent text-lg">{batch.remainingQuantity}</div>
                                <div className="text-[10px] text-muted tracking-widest font-bold">Bs. {batch.unitCost.toFixed(2)}/u</div>
                              </div>
                            </div>
                          ))}
                        {batches.filter(b => b.productId === product.id && b.remainingQuantity > 0).length === 0 && (
                          <div className="text-sm text-muted italic text-center py-4 bg-black/5 dark:bg-white/5 rounded-xl">Sin existencias</div>
                        )}
                      </div>
                    </div>
                  </div>
                ))}

                {products.length === 0 && (
                  <div className="col-span-full py-16 px-8 text-center bg-sidebar rounded-[2.5rem] border-2 border-dashed border-main/30 flex flex-col items-center justify-center space-y-6">
                    <div className="w-16 h-16 rounded-3xl bg-accent/10 text-accent flex items-center justify-center">
                      <Warehouse size={32} />
                    </div>
                    <div>
                      <h3 className="text-xl font-black uppercase tracking-tight text-main mb-1">
                        Empresa actual sin productos registrados
                      </h3>
                      <p className="text-xs text-muted font-bold max-w-md mx-auto uppercase tracking-wider leading-relaxed">
                        Esta empresa no tiene productos guardados. Si guardó sus productos en otra empresa o cuenta, puede cambiar inmediatamente:
                      </p>
                    </div>

                    {/* Quick switch to existing companies with products */}
                    {availableCompanies.filter(c => c.productCount > 0 && c.id !== effectiveCompanyId).length > 0 && (
                      <div className="w-full max-w-md p-4 bg-accent/5 border border-accent/20 rounded-2xl text-left space-y-3">
                        <span className="text-[10px] font-black uppercase tracking-widest text-accent block">
                          Empresas con datos guardados disponibles:
                        </span>
                        <div className="space-y-2">
                          {availableCompanies
                            .filter(c => c.productCount > 0 && c.id !== effectiveCompanyId)
                            .map(c => (
                              <button
                                key={c.id}
                                onClick={() => handleSwitchCompany(c.id, c.name)}
                                className="w-full flex items-center justify-between p-3 rounded-xl bg-card border border-main hover:border-accent transition-all text-xs font-bold"
                              >
                                <span className="font-black uppercase tracking-wider text-main">{c.name}</span>
                                <span className="text-accent text-[10px] font-black uppercase">
                                  {c.productCount} {c.productCount === 1 ? 'producto' : 'productos'} • {c.batchCount} lotes →
                                </span>
                              </button>
                            ))}
                        </div>
                      </div>
                    )}

                    <div className="flex flex-wrap items-center justify-center gap-3">
                      <button
                        onClick={() => { 
                          setProductRecipe([]);
                          setProductPackaging({ packagingType: 'ninguno', unitsPerBox: 0, unitsPerBundle: 0 });
                          setShowModal('product'); 
                          setFormData({ quantity: 0, cost: 0, minStock: 0, supplier: '', name: '', description: '', category: '', unit: 'unidad', type: 'producto terminado', productDescription: '', entryDate: new Date().toISOString().split('T')[0] }); 
                        }}
                        className="flex items-center gap-2 px-6 py-3.5 bg-accent text-white rounded-xl font-black uppercase text-xs tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-lg shadow-brand-gold-shadow"
                      >
                        <Plus size={16} />
                        <span>Crear Producto</span>
                      </button>

                      <button
                        onClick={handleSeedSystemData}
                        disabled={isSeedingData}
                        className="flex items-center gap-2 px-5 py-3.5 bg-sidebar border border-muted/30 text-muted hover:text-accent hover:border-accent/40 rounded-xl font-bold uppercase text-xs tracking-widest transition-all disabled:opacity-50"
                      >
                        <Sparkles size={16} className={isSeedingData ? 'animate-spin' : ''} />
                        <span>{isSeedingData ? 'Cargando Demo...' : 'Cargar Catálogo Demo'}</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'purchases' && (
            <motion.div 
                key="purchases"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
            >
                <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8">
                    <div>
                        <h2 className="text-2xl md:text-5xl font-black tracking-tight text-main mb-4">Módulo de Compras</h2>
                        <div className="flex gap-6 text-[10px] font-black uppercase tracking-[0.2em]">
                            <div className="flex items-center gap-2">
                                <span className="text-muted">Gestión de</span>
                                <span className="text-accent">Adquisiciones</span>
                            </div>
                        </div>
                    </div>
                    
                    <button 
                      onClick={() => {
                        const headers = ['Fecha', 'Producto', 'Proveedor', 'Cantidad', 'Costo Unit.', 'Total', 'Realizado por'];
                        const rows = filteredPurchaseReports.map(b => {
                          const product = products.find(p => p.id === b.productId);
                          return [
                            format(new Date(b.entryDate), 'dd/MM/yyyy HH:mm'),
                            product?.name || 'N/A',
                            b.supplier,
                            b.quantity,
                            b.unitCost.toFixed(2),
                            (b.quantity * b.unitCost).toFixed(2),
                            b.performedBy || 'Sistema'
                          ];
                        });
                        const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                        const workbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(workbook, worksheet, 'Compras');
                        XLSX.writeFile(workbook, `reporte_compras_${format(new Date(), 'yyyyMMdd')}.xlsx`);
                      }}
                      className="flex items-center gap-2 bg-sidebar border border-main px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-accent hover:text-white hover:border-accent transition-all shadow-sm active:scale-95"
                    >
                      <Download size={14} />
                      Excel Reporte
                    </button>
                </header>

                {/* Report Controls */}
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm space-y-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-2">
                    <div className="flex items-center gap-2">
                      <Settings size={16} className="text-accent" />
                      <span className="text-[10px] font-black uppercase tracking-[0.2em] text-muted">Filtros de Reporte de Compras</span>
                    </div>
                    <div className="flex flex-wrap items-center gap-3">
                      {selectedPurchases.length === 1 && (
                        <button 
                          onClick={() => {
                            const first = batches.find(b => b.id === selectedPurchases[0]);
                            if (first) {
                              const day = format(new Date(first.entryDate), 'yyyy-MM-dd');
                              const sameDayIds = filteredPurchaseReports
                                .filter(b => b.productId === first.productId && format(new Date(b.entryDate), 'yyyy-MM-dd') === day)
                                .map(b => b.id);
                              setSelectedPurchases(sameDayIds);
                            }
                          }}
                          className="bg-accent/10 text-accent border border-accent/20 px-6 py-3 rounded-xl text-[10px] font-black uppercase tracking-[0.2em] active:scale-95 transition-all hover:bg-accent/20"
                        >
                          Auto-seleccionar misma fecha
                        </button>
                      )}
                      {selectedPurchases.length > 1 && (
                        <button 
                          onClick={() => handleConsolidatePurchases()}
                          className="bg-accent text-white px-6 py-3 rounded-xl text-[10px] font-black uppercase tracking-[0.2em] shadow-lg shadow-brand-gold-shadow animate-bounce active:scale-95 transition-all"
                        >
                          Consolidar Seleccionados ({selectedPurchases.length})
                        </button>
                      )}
                      {selectedPurchases.length > 0 && (
                        <button 
                          onClick={() => {
                            const selected = batches.filter(b => selectedPurchases.includes(b.id));
                            setBulkPurchasesToDelete(selected);
                          }}
                          className="bg-red-500/10 text-red-500 border border-red-500/20 px-6 py-3 rounded-xl text-[10px] font-black uppercase tracking-[0.2em] hover:bg-red-500 hover:text-white transition-all active:scale-95 flex items-center gap-1.5"
                          title="Eliminar todos los registros de compra seleccionados"
                        >
                          <Trash2 size={13} />
                          Eliminar Seleccionados ({selectedPurchases.length})
                        </button>
                      )}
                    </div>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Buscar Proveedor o Producto</label>
                      <input 
                        type="text"
                        placeholder="Nombre..."
                        value={purchaseFilterSupplier}
                        onChange={(e) => setPurchaseFilterSupplier(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Fecha Inicial</label>
                      <input 
                        type="date"
                        value={purchaseFilterStartDate}
                        onChange={(e) => setPurchaseFilterStartDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Fecha Final</label>
                      <input 
                        type="date"
                        value={purchaseFilterEndDate}
                        onChange={(e) => setPurchaseFilterEndDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>
                  </div>
                </div>

                {/* Purchase List Table */}
                <div className="bg-sidebar rounded-3xl border border-main overflow-hidden shadow-2xl relative">
                  <div className="px-8 py-6 border-b border-main flex justify-between items-center bg-black/10 dark:bg-white/5">
                    <div className="flex items-center gap-4">
                      <div className="w-2 h-2 rounded-full bg-accent animate-pulse" />
                      <span className="text-xs font-black tracking-[0.4em] text-muted uppercase">Historial de Compras Realizadas</span>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left">
                      <thead className="text-muted text-[10px] font-bold uppercase tracking-wider">
                        <tr>
                          <th className="px-8 py-5 border-b border-main w-10">
                            <input 
                              type="checkbox" 
                              className="accent-accent"
                              checked={selectedPurchases.length === filteredPurchaseReports.length && filteredPurchaseReports.length > 0}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedPurchases(filteredPurchaseReports.map(b => b.id));
                                } else {
                                  setSelectedPurchases([]);
                                }
                              }}
                            />
                          </th>
                          <th className="px-8 py-5 border-b border-main">Fecha</th>
                          <th className="px-8 py-5 border-b border-main">Producto</th>
                          <th className="px-8 py-5 border-b border-main">Categoría</th>
                          <th className="px-8 py-5 border-b border-main">Proveedor</th>
                          <th className="px-8 py-5 border-b border-main">Ctd.</th>
                          <th className="px-8 py-5 border-b border-main text-right">Costo Unit.</th>
                          <th className="px-8 py-5 border-b border-main text-right">Total</th>
                          <th className="px-8 py-5 border-b border-main">Usuario</th>
                          <th className="px-6 py-5 border-b border-main text-center">Acciones</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-main/5">
                        {filteredPurchaseReports.length > 0 ? (
                          filteredPurchaseReports.map((buy) => {
                            const product = products.find(p => p.id === buy.productId);
                            return (
                              <tr key={buy.id} className={cn(
                                "hover:bg-accent/5 transition-all group",
                                selectedPurchases.includes(buy.id) && "bg-accent/10"
                              )}>
                                <td className="px-8 py-5 border-b border-main/5">
                                  <input 
                                    type="checkbox" 
                                    className="accent-accent"
                                    checked={selectedPurchases.includes(buy.id)}
                                    onChange={() => {
                                      if (selectedPurchases.includes(buy.id)) {
                                        setSelectedPurchases(selectedPurchases.filter(id => id !== buy.id));
                                      } else {
                                        setSelectedPurchases([...selectedPurchases, buy.id]);
                                      }
                                    }}
                                  />
                                </td>
                                <td className="px-8 py-5 whitespace-nowrap text-muted font-medium">{format(new Date(buy.entryDate), 'dd/MM/yy HH:mm')}</td>
                                <td className="px-8 py-5 text-main font-black whitespace-nowrap uppercase tracking-tight">{product?.name || 'Inexistente'}</td>
                                <td className="px-8 py-5 text-muted font-bold text-[11px] opacity-60 italic">{product?.category || 'N/A'}</td>
                                <td className="px-8 py-5 text-main font-bold whitespace-nowrap">{buy.supplier}</td>
                                <td className="px-8 py-5 text-accent font-black">
                                  <div>{buy.quantity} <span className="text-[9px] opacity-50">{product?.unit}</span></div>
                                  {buy.remainingQuantity < buy.quantity && (
                                    <span className="text-[8px] font-bold text-muted block opacity-80" title="Cantidad restante disponible en almacén">
                                      Disp: {buy.remainingQuantity}
                                    </span>
                                  )}
                                </td>
                                <td className="px-8 py-5 text-right font-mono text-xs font-bold text-muted">Bs. {buy.unitCost.toLocaleString()}</td>
                                <td className="px-8 py-5 text-right font-mono text-xs font-black text-main">Bs. {(buy.quantity * buy.unitCost).toLocaleString()}</td>
                                <td className="px-8 py-5 text-muted font-medium text-[10px] opacity-70 italic">{buy.performedBy || 'Sistema'}</td>
                                <td className="px-6 py-5 border-b border-main/5 text-center whitespace-nowrap">
                                  <div className="flex items-center justify-center gap-1">
                                    <button
                                      onClick={() => handleOpenEditPurchase(buy)}
                                      className="p-2 text-accent hover:bg-accent/15 rounded-xl transition-all"
                                      title="Editar Compra (Fecha, Proveedor, Cantidad, Costo, Notas)"
                                    >
                                      <Pencil size={15} />
                                    </button>
                                    <button
                                      onClick={() => handleDeletePurchase(buy)}
                                      className="p-2 text-red-500 hover:bg-red-500/15 rounded-xl transition-all"
                                      title="Eliminar Compra"
                                    >
                                      <Trash2 size={15} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan={10} className="px-8 py-20 text-center text-muted font-black uppercase tracking-[0.2em] opacity-40">No se encontraron registros de compras</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="pt-10">
                    <div className="flex items-center gap-4 mb-6">
                        <Plus size={20} className="text-accent" />
                        <h2 className="text-xl font-black uppercase tracking-tighter">Registrar Nueva Adquisición</h2>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                        {products.map(product => (
                            <div key={product.id} className="bg-sidebar p-6 rounded-3xl border border-main hover:border-accent transition-all shadow-sm">
                                <h3 className="text-lg font-black uppercase mb-1">{product.name}</h3>
                                <p className="text-xs text-muted mb-4">{product.category}</p>
                                <div className="flex items-center justify-between mb-6">
                                    <span className="text-xs font-bold text-muted uppercase tracking-widest">Stock actual</span>
                                    <span className="text-sm font-black text-accent">{productStock[product.id]} {product.unit}s</span>
                                </div>
                                <button 
                                    onClick={() => { setSelectedProductId(product.id); setShowModal('purchase'); setFormData({ ...formData, productDescription: '' }); }}
                                    className="w-full bg-accent text-white py-4 rounded-2xl font-black uppercase tracking-[0.1em] text-[10px] shadow-lg shadow-brand-gold-shadow active:scale-95 transition-all"
                                >
                                    Registrar Compra
                                </button>
                            </div>
                        ))}
                    </div>
                </div>
            </motion.div>
          )}

          {activeTab === 'sales' && (
            <motion.div 
                key="sales"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
            >
                <header className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-sidebar p-6 rounded-2xl border border-main shadow-sm mb-8">
                    <div>
                        <h2 className="text-3xl font-black uppercase tracking-tighter">Módulo de Ventas</h2>
                        <p className="text-muted text-sm font-bold opacity-60 uppercase tracking-widest">
                          Registro, Historial de Ventas, Edición, Eliminación y Verificación PEPS
                        </p>
                    </div>
                    <div className="flex flex-wrap items-center gap-3">
                      <button 
                        onClick={handleExportSalesXLSX}
                        className="flex items-center gap-2 bg-sidebar border border-main px-5 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-accent hover:text-white hover:border-accent transition-all shadow-sm active:scale-95"
                        title="Exportar reporte de ventas a Excel"
                      >
                        <Download size={14} />
                        Excel Ventas
                      </button>
                      <button
                        onClick={() => {
                          setSelectedProductId(null);
                          setSaleItems([]);
                          setFormData({
                            quantity: 0,
                            cost: 0,
                            supplier: '',
                            name: '',
                            description: '',
                            category: '',
                            unit: 'unidad',
                            type: 'producto terminado',
                            productDescription: 'Venta comercial',
                            entryDate: new Date().toISOString().split('T')[0]
                          });
                          setShowModal('sale');
                        }}
                        className="bg-accent text-white px-5 py-3 rounded-xl font-black text-xs uppercase tracking-widest hover:brightness-110 transition-all flex items-center gap-2 shadow-lg shadow-brand-gold-shadow"
                      >
                        <ShoppingCart size={16} />
                        Nueva Venta / Venta Masiva
                      </button>
                    </div>
                </header>

                {/* Sales Summary KPI Cards */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
                  <div className="bg-sidebar p-5 rounded-2xl border border-main shadow-sm">
                    <span className="text-[10px] font-black uppercase tracking-widest text-muted block mb-1">
                      Total Ventas (Bs.)
                    </span>
                    <span className="text-2xl font-black text-accent block">
                      Bs. {salesSummary.totalRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                    <span className="text-[9px] font-bold text-muted uppercase tracking-wider block mt-1">
                      {filteredSaleReports.length} {filteredSaleReports.length === 1 ? 'operación' : 'operaciones'}
                    </span>
                  </div>

                  <div className="bg-sidebar p-5 rounded-2xl border border-main shadow-sm">
                    <span className="text-[10px] font-black uppercase tracking-widest text-muted block mb-1">
                      Unidades Vendidas
                    </span>
                    <span className="text-2xl font-black text-main block">
                      {salesSummary.totalUnits.toLocaleString()}
                    </span>
                    <span className="text-[9px] font-bold text-muted uppercase tracking-wider block mt-1">
                      En el período filtrado
                    </span>
                  </div>

                  <div className="bg-sidebar p-5 rounded-2xl border border-main shadow-sm">
                    <span className="text-[10px] font-black uppercase tracking-widest text-muted block mb-1">
                      Costo de Ventas (COGS)
                    </span>
                    <span className="text-2xl font-black text-muted block font-mono">
                      Bs. {salesSummary.totalCogs.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                    <span className="text-[9px] font-bold text-muted uppercase tracking-wider block mt-1">
                      Valuación PEPS
                    </span>
                  </div>

                  <div className="bg-sidebar p-5 rounded-2xl border border-main shadow-sm">
                    <span className="text-[10px] font-black uppercase tracking-widest text-muted block mb-1">
                      Utilidad Bruta Estimada
                    </span>
                    <span className={cn(
                      "text-2xl font-black block",
                      salesSummary.grossProfit >= 0 ? "text-emerald-500" : "text-red-500"
                    )}>
                      Bs. {salesSummary.grossProfit.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                    <span className="text-[9px] font-bold text-emerald-500 uppercase tracking-wider block mt-1">
                      Margen: {salesSummary.profitMargin.toFixed(1)}%
                    </span>
                  </div>
                </div>

                {/* Sales Filter Bar */}
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm space-y-6 mb-8">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <Settings size={16} className="text-accent" />
                      <span className="text-[10px] font-black uppercase tracking-[0.2em] text-muted">
                        Filtros de Registro de Ventas
                      </span>
                    </div>
                    {filteredSaleReports.length > 0 && (
                      <span className="text-[10px] font-black uppercase tracking-wider text-accent bg-accent/10 px-3 py-1 rounded-full">
                        {filteredSaleReports.length} {filteredSaleReports.length === 1 ? 'Venta encontrada' : 'Ventas encontradas'}
                      </span>
                    )}
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                        Buscar Producto, Factura o Cliente
                      </label>
                      <div className="relative">
                        <input 
                          type="text"
                          placeholder="Buscar por nombre, nota, factura..."
                          value={saleFilterSearch}
                          onChange={(e) => setSaleFilterSearch(e.target.value)}
                          className="w-full bg-main/5 border border-main/20 rounded-xl pl-10 pr-4 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                        />
                        <Search size={14} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-muted" />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                        Fecha Inicial
                      </label>
                      <input 
                        type="date"
                        value={saleFilterStartDate}
                        onChange={(e) => setSaleFilterStartDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">
                        Fecha Final
                      </label>
                      <input 
                        type="date"
                        value={saleFilterEndDate}
                        onChange={(e) => setSaleFilterEndDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>
                  </div>
                </div>

                {/* Sales Movements Table */}
                <div className="bg-sidebar rounded-3xl border border-main overflow-hidden shadow-2xl relative mb-10">
                  <div className="px-8 py-6 border-b border-main flex justify-between items-center bg-black/10 dark:bg-white/5">
                    <div className="flex items-center gap-4">
                      <div className="w-2 h-2 rounded-full bg-accent animate-pulse" />
                      <span className="text-xs font-black tracking-[0.4em] text-muted uppercase">
                        Registro y Movimientos de Ventas Realizadas
                      </span>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left">
                      <thead className="text-muted text-[10px] font-bold uppercase tracking-wider">
                        <tr>
                          <th className="px-6 py-5 border-b border-main">Fecha</th>
                          <th className="px-6 py-5 border-b border-main">Producto</th>
                          <th className="px-6 py-5 border-b border-main">Categoría</th>
                          <th className="px-6 py-5 border-b border-main text-center">Ctd.</th>
                          <th className="px-6 py-5 border-b border-main text-right">P. Venta Unit.</th>
                          <th className="px-6 py-5 border-b border-main text-right">Total Venta</th>
                          <th className="px-6 py-5 border-b border-main text-right">Costo PEPS</th>
                          <th className="px-6 py-5 border-b border-main text-right">Utilidad</th>
                          <th className="px-6 py-5 border-b border-main">Factura / Referencia</th>
                          <th className="px-6 py-5 border-b border-main">Vendedor</th>
                          <th className="px-6 py-5 border-b border-main text-center">Acciones</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-main/5">
                        {filteredSaleReports.length > 0 ? (
                          filteredSaleReports.map((sale) => {
                            const product = products.find(p => p.id === sale.productId);
                            const unitPrice = sale.salePrice || sale.unitCost || 0;
                            const totalRevenue = sale.quantity * unitPrice;
                            const totalCost = sale.quantity * sale.unitCost;
                            const profit = totalRevenue - totalCost;

                            return (
                              <tr key={sale.id} className="hover:bg-accent/5 transition-all group">
                                <td className="px-6 py-5 whitespace-nowrap text-muted font-medium text-xs">
                                  {format(new Date(sale.date), 'dd/MM/yy HH:mm')}
                                </td>
                                <td className="px-6 py-5 text-main font-black whitespace-nowrap uppercase tracking-tight text-xs">
                                  {product?.name || 'Inexistente'}
                                </td>
                                <td className="px-6 py-5 text-muted font-bold text-[11px] opacity-60 italic">
                                  {product?.category || 'N/A'}
                                </td>
                                <td className="px-6 py-5 text-accent font-black text-center text-xs">
                                  {sale.quantity} <span className="text-[9px] opacity-50">{product?.unit}</span>
                                </td>
                                <td className="px-6 py-5 text-right font-mono text-xs font-bold text-main">
                                  Bs. {unitPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="px-6 py-5 text-right font-mono text-xs font-black text-accent">
                                  Bs. {totalRevenue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="px-6 py-5 text-right font-mono text-xs font-medium text-muted">
                                  Bs. {sale.unitCost.toFixed(2)}
                                </td>
                                <td className={cn(
                                  "px-6 py-5 text-right font-mono text-xs font-black",
                                  profit >= 0 ? "text-emerald-500" : "text-red-500"
                                )}>
                                  Bs. {profit.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                                <td className="px-6 py-5 text-muted font-medium text-[11px] max-w-[180px] truncate" title={sale.reference || sale.description || ''}>
                                  {sale.reference || sale.description || 'Sin notas'}
                                </td>
                                <td className="px-6 py-5 text-muted font-medium text-[10px] opacity-70 italic">
                                  {sale.performedBy || 'Sistema'}
                                </td>
                                <td className="px-6 py-5 text-center whitespace-nowrap">
                                  <div className="flex items-center justify-center gap-1">
                                    <button
                                      onClick={() => handleOpenEditSale(sale)}
                                      className="p-2 text-accent hover:bg-accent/15 rounded-xl transition-all"
                                      title="Editar Venta (Precio, Cantidad, Fecha, Factura/Notas)"
                                    >
                                      <Pencil size={15} />
                                    </button>
                                    <button
                                      onClick={() => handleDeleteSale(sale)}
                                      className="p-2 text-red-500 hover:bg-red-500/15 rounded-xl transition-all"
                                      title="Eliminar Venta y Restituir Stock"
                                    >
                                      <Trash2 size={15} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan={11} className="px-8 py-20 text-center text-muted font-black uppercase tracking-[0.2em] opacity-40">
                              No se encontraron registros de ventas
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {salesTrendData.length > 0 && (
                  <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm mb-8">
                    <div className="flex justify-between items-center mb-8">
                      <div>
                        <h3 className="text-xs font-black uppercase tracking-[0.3em] text-muted mb-1">Tendencias de Ventas</h3>
                        <p className="text-sm font-black text-main uppercase tracking-tighter">Histórico de Salidas (Bs.)</p>
                      </div>
                      <div className="flex items-center gap-4">
                        <div className="flex items-center gap-2">
                          <div className="w-3 h-3 rounded-full bg-accent" />
                          <span className="text-[10px] font-black uppercase text-muted tracking-widest">Ingresos Totales</span>
                        </div>
                      </div>
                    </div>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={salesTrendData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorAmount" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="var(--color-accent)" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="var(--color-accent)" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--color-main-20)" opacity={0.1} />
                          <XAxis 
                            dataKey="formattedDate" 
                            axisLine={false} 
                            tickLine={false} 
                            tick={{ fontSize: 10, fontWeight: 900, fill: 'var(--color-muted)' }}
                            dy={10}
                          />
                          <YAxis 
                            axisLine={false} 
                            tickLine={false} 
                            tick={{ fontSize: 10, fontWeight: 900, fill: 'var(--color-muted)' }}
                            tickFormatter={(value) => `Bs. ${value.toLocaleString()}`}
                          />
                          <RechartsTooltip 
                            contentStyle={{ 
                              backgroundColor: 'var(--color-sidebar)', 
                              border: '1px solid var(--color-main)', 
                              borderRadius: '1rem',
                              fontSize: '10px',
                              fontWeight: 900,
                              textTransform: 'uppercase',
                              letterSpacing: '0.1em'
                            }}
                            itemStyle={{ color: 'var(--color-accent)' }}
                          />
                          <Area 
                            type="monotone" 
                            dataKey="amount" 
                            stroke="var(--color-accent)" 
                            strokeWidth={4}
                            fillOpacity={1} 
                            fill="url(#colorAmount)" 
                          />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    {products.map(product => (
                        <div key={product.id} className="bg-sidebar p-6 rounded-3xl border border-main hover:border-accent transition-all shadow-sm">
                            <h3 className="text-lg font-black uppercase mb-1">{product.name}</h3>
                            <p className="text-xs text-muted mb-4">{product.category}</p>
                            <div className="flex items-center justify-between mb-6">
                                <span className="text-xs font-bold text-muted uppercase tracking-widest">Disponibilidad</span>
                                <span className="text-sm font-black text-accent">{productStock[product.id]} {product.unit}s</span>
                            </div>
                            <button 
                                disabled={productStock[product.id] <= 0}
                                onClick={() => { setSelectedProductId(product.id); setShowModal('sale'); setFormData({ ...formData, productDescription: '' }); }}
                                className={cn(
                                    "w-full py-4 rounded-2xl font-black uppercase tracking-[0.1em] text-[10px] transition-all",
                                    productStock[product.id] > 0 
                                      ? "bg-accent text-white shadow-lg shadow-brand-gold-shadow active:scale-95" 
                                      : "bg-main/10 text-muted cursor-not-allowed opacity-50"
                                )}
                            >
                                Registrar Venta
                            </button>
                        </div>
                    ))}
                </div>
            </motion.div>
          )}

          {activeTab === 'kardex' && (
            <motion.div 
              key="kardex"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-10"
            >
              <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                <div>
                  <div className="flex items-center gap-3 mb-2">
                    <h2 className="text-2xl md:text-5xl font-black tracking-tight text-main">Kardex</h2>
                    <span className="bg-emerald-500/10 text-emerald-500 border border-emerald-500/30 text-[10px] md:text-xs font-black uppercase px-3 py-1 rounded-full tracking-wider flex items-center gap-1.5 shadow-sm">
                      <PackageCheck size={14} /> Solo Productos Terminados
                    </span>
                  </div>
                  <p className="text-muted text-xs font-bold uppercase tracking-wider mb-3">
                    Control y Valoración Contable al Costo Promedio al momento de salida exclusivo para Almacén de Producto Terminado
                  </p>
                  <div className="flex flex-wrap gap-4 md:gap-6 text-[10px] font-black uppercase tracking-[0.2em]">
                      <div className="flex items-center gap-2">
                          <span className="text-muted">Método:</span>
                          <span className="text-accent">Costo Promedio Ponderado (al momento de salida)</span>
                      </div>
                      <span className="opacity-20 text-main font-thin">|</span>
                      <div className="flex items-center gap-2">
                          <span className="text-muted">Enfoque:</span>
                          <span className="text-emerald-500">Producto Terminado</span>
                      </div>
                      <span className="opacity-20 text-main font-thin">|</span>
                      <div className="flex items-center gap-2">
                          <span className="text-muted">Documento:</span>
                          <span className="text-accent">Libro de Operaciones</span>
                      </div>
                  </div>
                </div>
                
                <div className="flex items-center gap-4">
                  <button 
                    onClick={() => setKardexPresentationMode(!kardexPresentationMode)}
                    className={cn(
                      "flex items-center gap-2 px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] transition-all shadow-sm active:scale-95 border",
                      kardexPresentationMode 
                        ? "bg-accent text-white border-accent" 
                        : "bg-sidebar border-main text-muted hover:text-main hover:border-accent"
                    )}
                  >
                    {kardexPresentationMode ? <EyeOff size={14} /> : <Eye size={14} />}
                    {kardexPresentationMode ? 'Modo Edición' : 'Modo Presentación'}
                  </button>

                  {!kardexPresentationMode && (
                    <div className="flex flex-wrap items-center gap-2">
                      <button 
                        onClick={() => handleExportXLSX('filtered')}
                        title="Exporta los movimientos de productos terminados según los filtros"
                        className="flex items-center gap-2 bg-sidebar border border-accent/40 text-accent px-5 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-accent hover:text-white hover:border-accent transition-all shadow-sm active:scale-95"
                      >
                        <FileSpreadsheet size={14} />
                        Exportar Excel (Filtros)
                      </button>
                      <button 
                        onClick={() => handleExportXLSX('all')}
                        title="Genera el libro contable de Tarjeta Kardex valuado al costo promedio de todos los productos terminados"
                        className="flex items-center gap-2 bg-sidebar border border-main px-4 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] text-muted hover:text-main hover:bg-main/5 transition-all shadow-sm active:scale-95"
                      >
                        <Download size={14} />
                        Exportar Todo P. Terminado
                      </button>
                    </div>
                  )}
                </div>
              </header>

              {/* Filtering Controls */}
              {!kardexPresentationMode && (
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm space-y-6">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <Settings size={16} className="text-accent" />
                      <span className="text-[10px] font-black uppercase tracking-[0.2em] text-muted">Filtros de Productos Terminados</span>
                    </div>
                    <span className="text-[10px] font-bold text-emerald-500 uppercase tracking-widest">
                      {products.filter(p => p.type === 'producto terminado').length} productos terminados registrados
                    </span>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                    {/* Search Field */}
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Búsqueda General</label>
                      <div className="relative">
                        <input 
                          type="text"
                          placeholder="Buscar por lote, referencia, responsable..."
                          value={kardexSearch}
                          onChange={(e) => setKardexSearch(e.target.value)}
                          className="w-full bg-main/5 border border-main/20 rounded-xl pl-12 pr-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                        />
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-muted/40" size={16} />
                      </div>
                    </div>

                    {/* Product Name Filter (Focused on Finished Products) */}
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1 flex items-center justify-between">
                        <span>Producto Terminado</span>
                        <span className="text-[9px] text-emerald-500 font-bold">Exclusivo</span>
                      </label>
                      <select 
                        value={kardexFilterProduct}
                        onChange={(e) => setKardexFilterProduct(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-4 py-3 text-xs font-bold focus:border-accent outline-none transition-all appearance-none text-main"
                      >
                        <option value="">TODOS LOS PRODUCTOS TERMINADOS</option>
                        {products
                          .filter(p => p.type === 'producto terminado')
                          .map(p => (
                            <option key={p.id} value={p.id}>
                              {p.name} ({p.unit}s) - Stock: {productStock[p.id] || 0}
                            </option>
                          ))
                        }
                      </select>
                    </div>

                    {/* Start Date Filter */}
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Fecha Desde</label>
                      <input 
                        type="date"
                        value={kardexFilterStartDate}
                        onChange={(e) => setKardexFilterStartDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>

                    {/* End Date Filter */}
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Fecha Hasta</label>
                      <input 
                        type="date"
                        value={kardexFilterEndDate}
                        onChange={(e) => setKardexFilterEndDate(e.target.value)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                      />
                    </div>

                    {/* Transaction Type Filter */}
                    <div className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-muted ml-1">Tipo de Operación</label>
                      <select 
                        value={kardexFilterType}
                        onChange={(e) => setKardexFilterType(e.target.value as any)}
                        className="w-full bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all appearance-none"
                      >
                        <option value="ALL">TODOS LOS MOVIMIENTOS</option>
                        <option value="ENTRY">ENTRADAS MANUALES</option>
                        <option value="EXIT">SALIDAS MANUALES</option>
                        <option value="PURCHASE">COMPRAS REGISTRADAS</option>
                        <option value="SALE">VENTAS REALIZADAS</option>
                      </select>
                    </div>
                  </div>

                  <div className="flex justify-end pt-2">
                    {(kardexFilterProduct || kardexFilterStartDate || kardexFilterEndDate || kardexFilterType !== 'ALL' || kardexSearch) && (
                      <button 
                        onClick={() => {
                          setKardexFilterProduct('');
                          setKardexFilterStartDate('');
                          setKardexFilterEndDate('');
                          setKardexFilterType('ALL');
                          setKardexSearch('');
                        }}
                        className="text-[10px] font-black uppercase tracking-widest text-accent hover:underline flex items-center gap-1 transition-all"
                      >
                        Limpiar todos los filtros
                      </button>
                    )}
                  </div>
                </div>
              )}

              <div className="bg-sidebar rounded-3xl border border-main overflow-hidden shadow-2xl relative">
                <div className="px-6 py-8 md:px-8 md:py-6 border-b border-main flex flex-col xl:flex-row justify-between items-start xl:items-center bg-black/10 dark:bg-white/5 gap-8">
                  <div className="flex items-center gap-3">
                    <ClipboardList size={18} className="text-accent" />
                    <div>
                      <span className="text-[10px] md:text-xs font-black tracking-[0.2em] md:tracking-[0.4em] text-muted uppercase block">
                        Consolidado Tarjeta Kardex • Costo Promedio al momento de salida
                      </span>
                      <span className="text-[9px] font-bold text-emerald-500 uppercase tracking-widest">
                        Almacén de Producto Terminado
                      </span>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-4 md:gap-8 w-full xl:w-auto">
                    <div className="flex flex-col items-start xl:items-end">
                      <span className="text-[8px] md:text-[9px] font-black uppercase text-emerald-500 tracking-wider">Valor Almacén P. Terminado</span>
                      <span className="text-xs md:text-sm font-black text-emerald-500">
                        Bs. {batches
                          .filter(b => products.some(p => p.id === b.productId && p.type === 'producto terminado') && b.remainingQuantity > 0)
                          .reduce((acc, b) => acc + (b.remainingQuantity * b.unitCost), 0)
                          .toLocaleString()}
                      </span>
                    </div>
                    <div className="flex flex-col items-start xl:items-end">
                      <span className="text-[8px] md:text-[9px] font-black uppercase text-muted/60 tracking-wider">Total Entradas</span>
                      <span className="text-xs md:text-sm font-black text-green-500">Bs. {filteredKardex.reduce((acc, curr) => acc + curr.entryTotal, 0).toLocaleString()}</span>
                    </div>
                    <div className="flex flex-col items-start xl:items-end">
                      <span className="text-[8px] md:text-[9px] font-black uppercase text-muted/60 tracking-wider">Costo Ventas</span>
                      <span className="text-xs md:text-sm font-black text-red-500">Bs. {filteredKardex.reduce((acc, curr) => acc + (curr.type === 'SALE' ? curr.exitTotal : 0), 0).toLocaleString()}</span>
                    </div>
                    <div className="flex flex-col items-start xl:items-end">
                      <span className="text-[8px] md:text-[9px] font-black uppercase text-muted/60 tracking-wider">Mermas / Ajustes</span>
                      <span className="text-xs md:text-sm font-black text-orange-500">Bs. {filteredKardex.reduce((acc, curr) => acc + (curr.type === 'EXIT' ? curr.exitTotal : 0), 0).toLocaleString()}</span>
                    </div>
                    <span className="bg-accent text-white text-[9px] md:text-[10px] font-black px-4 py-1.5 rounded-full tracking-widest uppercase shadow-lg shadow-brand-gold-shadow ml-auto xl:ml-0">
                    {Math.min(kardexCurrentPage * KARDEX_ITEMS_PER_PAGE, filteredKardex.length)} de {filteredKardex.length}
                    </span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                    <table className="w-full text-sm text-left min-w-[1300px]">
                    <thead className="text-[10px] font-black uppercase tracking-widest text-muted bg-main/5 border-b border-main/10">
                        <tr>
                        <th className="px-8 py-6" rowSpan={2}>Fecha/Hora</th>
                        <th className="px-8 py-6" rowSpan={2}>Producto Terminado</th>
                        <th className="px-8 py-6" rowSpan={2}>Referencia</th>
                        <th className="px-8 py-6 text-center border-l border-main/10 bg-black/10 hover:bg-black/20 transition-colors" colSpan={3}>Operación Entrada (Producción/Altas)</th>
                        <th className="px-8 py-6 text-center border-l border-main/10" colSpan={4}>Operación Salida (Costo de Ventas al Costo Promedio)</th>
                        <th className="px-8 py-6 text-center border-l border-main/10" rowSpan={2}>Responsable</th>
                        <th className="px-8 py-6 text-center border-l border-main/10 bg-black/10 hover:bg-black/20 transition-colors" colSpan={3}>Saldos (del Producto)</th>
                        </tr>
                        <tr className="border-t border-main/10 bg-main/5">
                        <th className="px-8 py-3 border-l border-main/10 text-center opacity-70">Cant.</th>
                        <th className="px-8 py-3 text-center opacity-70">Cost.U</th>
                        <th className="px-8 py-3 text-center opacity-70">Val.Tot</th>
                        <th className="px-8 py-3 border-l border-main/10 text-center opacity-70">Cant.</th>
                        <th className="px-8 py-3 text-center opacity-70" title="Costo promedio ponderado al momento de salida">Cost.Prom</th>
                        <th className="px-8 py-3 text-center opacity-70">Val.Tot</th>
                        <th className="px-8 py-3 text-center opacity-70 border-l border-main/5 text-accent">P. Venta</th>
                        <th className="px-8 py-3 border-l border-main/10 text-center opacity-70">Cant.</th>
                        <th className="px-8 py-3 text-center opacity-70">Prom.P</th>
                        <th className="px-8 py-3 text-center opacity-70">Patrimonio</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-main/5">
                        {paginatedKardex.length > 0 ? (
                          paginatedKardex.map((entry, idx) => {
                            const prod = products.find(p => p.id === entry.productId);
                            return (
                              <tr key={`${entry.date}-${idx}`} className="hover:bg-accent/5 transition-all group border-l-4 border-l-transparent hover:border-l-accent">
                                  <td className="px-8 py-5 whitespace-nowrap text-muted font-medium">{format(new Date(entry.date), 'dd/MM/yy HH:mm')}</td>
                                  <td className="px-8 py-5 text-main font-black whitespace-nowrap uppercase tracking-tight">
                                    <div>{entry.productName}</div>
                                    {prod?.packaging && prod.packaging.packagingType !== 'ninguno' && (
                                      <div className="flex items-center gap-1.5 mt-1 text-[9px] font-bold">
                                        {prod.packaging.unitsPerBox && prod.packaging.unitsPerBox > 0 && (
                                          <span className="bg-blue-500/10 text-blue-600 dark:text-blue-400 px-1.5 py-0.5 rounded">
                                            📦 {prod.packaging.unitsPerBox} un/caja
                                          </span>
                                        )}
                                        {prod.packaging.unitsPerBundle && prod.packaging.unitsPerBundle > 0 && (
                                          <span className="bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 px-1.5 py-0.5 rounded">
                                            🧺 {prod.packaging.unitsPerBundle} un/bulto
                                          </span>
                                        )}
                                      </div>
                                    )}
                                  </td>
                                  <td className="px-8 py-5 text-muted font-bold text-[11px] opacity-60 italic whitespace-nowrap">
                                    <div className="flex items-center gap-3">
                                      <span>{entry.reference}</span>
                                      {entry.batchId && (
                                        <button 
                                          onClick={() => {
                                            setSelectedBatchId(entry.batchId!);
                                            const batch = batches.find(b => b.id === entry.batchId);
                                            if (batch) {
                                              setFormData({
                                                ...formData,
                                                quantity: batch.quantity,
                                                cost: batch.unitCost,
                                                supplier: batch.supplier,
                                                productDescription: batch.description || '',
                                                entryDate: batch.entryDate ? (batch.entryDate.includes('T') ? batch.entryDate.split('T')[0] : batch.entryDate) : new Date().toISOString().split('T')[0]
                                              });
                                              setShowModal('edit_batch');
                                            }
                                          }}
                                          className="p-1.5 hover:bg-accent/20 text-accent rounded-lg transition-all opacity-0 group-hover:opacity-100 flex items-center gap-1 border border-accent/10 bg-accent/5 backdrop-blur-sm"
                                          title="Configurar Origen de Lote"
                                        >
                                          <Settings size={10} />
                                          <span className="text-[8px] font-black uppercase">Config.</span>
                                        </button>
                                      )}
                                    </div>
                                  </td>
                                  {/* Entradas */}
                                  <td className="px-8 py-5 border-l border-main/10 text-green-500 font-black text-center text-sm">
                                    {entry.entryQty > 0 ? (
                                      <div>
                                        <div>+{entry.entryQty}</div>
                                        {prod?.packaging?.unitsPerBox && prod.packaging.unitsPerBox > 0 && (
                                          <span className="text-[8px] font-bold text-muted block opacity-80">
                                            ({(entry.entryQty / prod.packaging.unitsPerBox).toFixed(1)} cajas)
                                          </span>
                                        )}
                                      </div>
                                    ) : '-'}
                                  </td>
                                  <td className="px-8 py-5 text-main text-center font-mono opacity-80 text-xs">{entry.entryUnitCost > 0 ? `Bs. ${entry.entryUnitCost.toFixed(2)}` : ''}</td>
                                  <td className="px-8 py-5 text-green-500 font-black text-center text-sm">{entry.entryTotal > 0 ? `Bs. ${entry.entryTotal.toFixed(2)}` : ''}</td>
                                  {/* Salidas */}
                                  <td className="px-8 py-5 border-l border-main/10 text-red-500 font-black text-center text-sm">
                                    {entry.exitQty > 0 ? (
                                      <div>
                                        <div>-{entry.exitQty}</div>
                                        {prod?.packaging?.unitsPerBox && prod.packaging.unitsPerBox > 0 && (
                                          <span className="text-[8px] font-bold text-muted block opacity-80">
                                            ({(entry.exitQty / prod.packaging.unitsPerBox).toFixed(1)} cajas)
                                          </span>
                                        )}
                                      </div>
                                    ) : '-'}
                                  </td>
                                  <td className="px-8 py-5 text-main text-center font-mono opacity-80 text-xs">{entry.exitUnitCost > 0 ? `Bs. ${entry.exitUnitCost.toFixed(2)}` : ''}</td>
                                  <td className="px-8 py-5 text-red-500 font-black text-center text-sm">{entry.exitTotal > 0 ? `Bs. ${entry.exitTotal.toFixed(2)}` : ''}</td>
                                  <td className="px-8 py-5 text-accent font-black text-center text-sm border-l border-main/5">
                                    {entry.salePrice ? `Bs. ${entry.salePrice.toLocaleString()}` : '-'}
                                  </td>
                                  {/* Audit */}
                                  <td className="px-8 py-5 border-l border-main/10 text-muted text-xs italic font-medium whitespace-nowrap">
                                    {entry.performedBy || 'Sistema'}
                                  </td>
                                  {/* Saldos */}
                                  <td className="px-8 py-5 border-l border-main/10 font-black text-accent text-center text-sm">
                                    <div>{entry.balanceQty}</div>
                                    {prod?.packaging?.unitsPerBox && prod.packaging.unitsPerBox > 0 && (
                                      <span className="text-[8px] font-bold text-muted block opacity-80">
                                        ({(entry.balanceQty / prod.packaging.unitsPerBox).toFixed(1)} cajas)
                                      </span>
                                    )}
                                  </td>
                                  <td className="px-8 py-5 text-muted text-center font-mono opacity-40 italic font-medium text-xs">Bs. {entry.balanceUnitCost.toFixed(2)}</td>
                                  <td className="px-8 py-5 text-accent font-black text-center shadow-accent/5 text-sm">Bs. {entry.balanceTotal.toFixed(2)}</td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan={14} className="px-8 py-16 text-center">
                              {products.filter(p => p.type === 'producto terminado').length === 0 ? (
                                <div className="max-w-md mx-auto space-y-3">
                                  <div className="w-12 h-12 rounded-2xl bg-emerald-500/10 text-emerald-500 flex items-center justify-center mx-auto">
                                    <PackageCheck size={24} />
                                  </div>
                                  <h4 className="text-sm font-black uppercase text-main">No hay Productos Terminados Registrados</h4>
                                  <p className="text-xs text-muted font-medium">
                                    La Tarjeta Kardex está enfocada exclusivamente en Productos Terminados. Registra tu primer producto terminado indicando sus consumos y empaque (caja o bulto).
                                  </p>
                                  <button
                                    onClick={() => {
                                      setFormData({
                                        ...formData,
                                        type: 'producto terminado',
                                        name: '',
                                        description: '',
                                        category: '',
                                        minStock: 10
                                      });
                                      setShowModal('product');
                                    }}
                                    className="px-6 py-2.5 bg-accent text-white rounded-xl font-black text-xs uppercase tracking-wider shadow-lg shadow-brand-gold-shadow hover:brightness-110 transition-all inline-flex items-center gap-2"
                                  >
                                    <Plus size={16} /> Crear Producto Terminado
                                  </button>
                                </div>
                              ) : (
                                <div className="text-muted font-bold uppercase tracking-[0.2em] opacity-40 italic">
                                  No se encontraron registros de productos terminados bajo los filtros aplicados
                                </div>
                              )}
                            </td>
                          </tr>
                        )}
                    </tbody>
                    </table>
                </div>

                {/* Pagination Controls */}
                {totalKardexPages > 1 && (
                  <div className="px-8 py-6 bg-black/10 dark:bg-white/5 border-t border-main flex justify-between items-center">
                    <span className="text-[10px] font-black uppercase tracking-widest text-muted">
                      Página {kardexCurrentPage} de {totalKardexPages}
                    </span>
                    <div className="flex gap-2">
                       <button 
                        disabled={kardexCurrentPage === 1}
                        onClick={() => setKardexCurrentPage(prev => Math.max(1, prev - 1))}
                        className="p-2 rounded-lg border border-main text-muted hover:text-accent hover:border-accent disabled:opacity-30 disabled:cursor-not-allowed transition-all"
                       >
                         <ChevronLeft size={18} />
                       </button>
                       <div className="flex items-center gap-1">
                          {Array.from({ length: Math.min(5, totalKardexPages) }, (_, i) => {
                             let pageNum = kardexCurrentPage;
                             if (kardexCurrentPage <= 3) pageNum = i + 1;
                             else if (kardexCurrentPage >= totalKardexPages - 2) pageNum = totalKardexPages - 4 + i;
                             else pageNum = kardexCurrentPage - 2 + i;

                             if (pageNum < 1 || pageNum > totalKardexPages) return null;

                             return (
                               <button 
                                key={pageNum}
                                onClick={() => setKardexCurrentPage(pageNum)}
                                className={cn(
                                  "w-8 h-8 rounded-lg text-[10px] font-black transition-all",
                                  kardexCurrentPage === pageNum 
                                    ? "bg-accent text-white" 
                                    : "text-muted hover:bg-main/5"
                                )}
                               >
                                 {pageNum}
                               </button>
                             );
                          })}
                       </div>
                       <button 
                        disabled={kardexCurrentPage === totalKardexPages}
                        onClick={() => setKardexCurrentPage(prev => Math.min(totalKardexPages, prev + 1))}
                        className="p-2 rounded-lg border border-main text-muted hover:text-accent hover:border-accent disabled:opacity-30 disabled:cursor-not-allowed transition-all"
                       >
                         <ChevronRight size={18} />
                       </button>
                    </div>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'audit' && (
            <motion.div 
              key="audit"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8"
            >
              <header>
                <h2 className="text-2xl md:text-5xl font-black tracking-tight text-main mb-2">Registro de Auditoría</h2>
                <p className="text-muted text-xs uppercase tracking-widest font-bold">Trazabilidad completa de operaciones y cambios de sistema</p>
              </header>

              <div className="bg-sidebar p-6 rounded-2xl border border-main shadow-sm space-y-4">
                <div className="flex flex-col md:flex-row gap-4">
                  <div className="relative flex-1">
                    <input 
                      type="text"
                      placeholder="Buscar en logs..."
                      value={auditSearch}
                      onChange={(e) => setAuditSearch(e.target.value)}
                      className="w-full bg-main/5 border border-main/20 rounded-xl pl-12 pr-5 py-3 text-xs font-bold focus:border-accent outline-none transition-all"
                    />
                    <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-muted/40" size={16} />
                  </div>
                  <select 
                    value={auditFilterAction}
                    onChange={(e) => setAuditFilterAction(e.target.value as any)}
                    className="bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none appearance-none min-w-[150px]"
                  >
                    <option value="ALL">TODAS LAS ACCIONES</option>
                    <option value="CREATE">CREACIONES</option>
                    <option value="UPDATE">MODIFICACIONES</option>
                    <option value="DELETE">ELIMINACIONES</option>
                    <option value="SYSTEM">SISTEMA</option>
                  </select>
                  <select 
                    value={auditFilterEntity}
                    onChange={(e) => setAuditFilterEntity(e.target.value as any)}
                    className="bg-main/5 border border-main/20 rounded-xl px-5 py-3 text-xs font-bold focus:border-accent outline-none appearance-none min-w-[150px]"
                  >
                    <option value="ALL">TODAS LAS ENTIDADES</option>
                    <option value="PRODUCT">PRODUCTOS</option>
                    <option value="BATCH">LOTES</option>
                    <option value="TRANSACTION">TRANSACCIONES</option>
                    <option value="TEAM">EQUIPO</option>
                    <option value="CONFIG">CONFIGURACIÓN</option>
                  </select>
                </div>

                <div className="flex flex-col md:flex-row gap-4">
                  <div className="flex-1 flex items-center gap-3">
                    <span className="text-[10px] font-black uppercase text-muted tracking-widest min-w-[70px]">Desde:</span>
                    <input 
                      type="date"
                      value={auditFilterStartDate}
                      onChange={(e) => setAuditFilterStartDate(e.target.value)}
                      className="flex-1 bg-main/5 border border-main/20 rounded-xl px-4 py-2 text-[11px] font-bold focus:border-accent outline-none transition-all uppercase"
                    />
                  </div>
                  <div className="flex-1 flex items-center gap-3">
                    <span className="text-[10px] font-black uppercase text-muted tracking-widest min-w-[70px]">Hasta:</span>
                    <input 
                      type="date"
                      value={auditFilterEndDate}
                      onChange={(e) => setAuditFilterEndDate(e.target.value)}
                      className="flex-1 bg-main/5 border border-main/20 rounded-xl px-4 py-2 text-[11px] font-bold focus:border-accent outline-none transition-all uppercase"
                    />
                  </div>
                  <button 
                    onClick={() => {
                      setAuditFilterStartDate('');
                      setAuditFilterEndDate('');
                      setAuditFilterAction('ALL');
                      setAuditFilterEntity('ALL');
                      setAuditSearch('');
                    }}
                    className="bg-main/10 text-muted px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-main/20 transition-all border border-main/10"
                  >
                    LIMPIAR FILTROS
                  </button>
                </div>
              </div>

              {/* Summary Cards */}
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-6">
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm flex flex-col justify-center">
                  <div className="text-[10px] font-black uppercase text-muted tracking-[0.2em] mb-3">Registros Totales</div>
                  <div className="text-4xl font-black text-main">{auditStats.total}</div>
                </div>
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm border-l-4 border-l-green-500 flex flex-col justify-center">
                  <div className="text-[10px] font-black uppercase text-muted tracking-[0.2em] mb-3 text-green-500/80">Creaciones</div>
                  <div className="text-4xl font-black text-green-500">{auditStats.creations}</div>
                </div>
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm border-l-4 border-l-blue-500 flex flex-col justify-center">
                  <div className="text-[10px] font-black uppercase text-muted tracking-[0.2em] mb-3 text-blue-500/80">Modificaciones</div>
                  <div className="text-4xl font-black text-blue-500">{auditStats.updates}</div>
                </div>
                <div className="bg-sidebar p-8 rounded-[2rem] border border-main shadow-sm border-l-4 border-l-red-500 flex flex-col justify-center">
                  <div className="text-[10px] font-black uppercase text-muted tracking-[0.2em] mb-3 text-red-500/80">Eliminaciones</div>
                  <div className="text-4xl font-black text-red-500">{auditStats.deletions}</div>
                </div>
              </div>

              <div className="bg-sidebar rounded-3xl border border-main overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead className="text-[10px] font-black uppercase tracking-widest text-muted bg-main/5 border-b border-main/10">
                      <tr>
                        <th className="px-8 py-4">Fecha/Hora</th>
                        <th className="px-8 py-4">Usuario</th>
                        <th className="px-8 py-4">Acción</th>
                        <th className="px-8 py-4">Entidad</th>
                        <th className="px-8 py-4">Detalles</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-main/5">
                      {filteredAuditLogs.map((log) => (
                          <tr key={log.id} className="hover:bg-main/5 transition-colors">
                            <td className="px-8 py-4 whitespace-nowrap text-[11px] font-bold text-muted">
                              {format(new Date(log.timestamp), 'dd/MM/yy HH:mm:ss')}
                            </td>
                            <td className="px-8 py-4 whitespace-nowrap">
                              <div className="text-[11px] font-black uppercase">{log.userName}</div>
                              <div className="text-[9px] text-muted opacity-50 font-mono">{log.userId.slice(0, 8)}</div>
                            </td>
                            <td className="px-8 py-4 whitespace-nowrap">
                              <span className={cn(
                                "text-[9px] font-black uppercase px-2 py-1 rounded border",
                                log.action === 'CREATE' ? "bg-green-500/10 text-green-500 border-green-500/20" :
                                log.action === 'UPDATE' ? "bg-blue-500/10 text-blue-500 border-blue-500/20" :
                                log.action === 'DELETE' ? "bg-red-500/10 text-red-500 border-red-500/20" :
                                "bg-main/10 text-muted border-main/20"
                              )}>
                                {log.action}
                              </span>
                            </td>
                            <td className="px-8 py-4 whitespace-nowrap text-[10px] font-black text-accent uppercase tracking-tighter">
                              {log.entityType}
                            </td>
                            <td className="px-8 py-4">
                              <div className="text-xs font-bold text-main">{log.details}</div>
                              {(log.previousValue || log.newValue) && (
                                <button 
                                  onClick={() => {
                                    console.log('Log Entry:', log.id);
                                    if (log.previousValue) console.log('Previous Value:', JSON.parse(log.previousValue));
                                    if (log.newValue) console.log('New Value:', JSON.parse(log.newValue));
                                    alert('Detalles técnicos enviados a la consola del navegador.');
                                  }}
                                  className="text-[9px] font-black uppercase text-accent hover:underline mt-1"
                                >
                                  Ver comparativa técnica
                                </button>
                              )}
                            </td>
                          </tr>
                        ))}
                      {auditLogs.length === 0 && (
                        <tr>
                          <td colSpan={5} className="px-8 py-20 text-center text-muted font-bold uppercase tracking-widest opacity-40 italic">
                            No hay registros de auditoría disponibles
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'configuration' && (
            <motion.div 
              key="configuration"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="space-y-8 max-w-2xl"
            >
              <header className="mb-10">
                <h2 className="text-2xl md:text-5xl font-black tracking-tight text-main mb-4">Configuración General</h2>
                <div className="flex gap-6 text-[10px] font-black uppercase tracking-[0.2em] text-muted">
                    <div className="flex items-center gap-2">
                        <Settings size={14} className="text-accent" />
                        <span>Ajustes del</span>
                        <span className="text-accent">Sistema</span>
                    </div>
                </div>
              </header>

              <div className="bg-sidebar p-10 rounded-[2.5rem] border border-main shadow-2xl space-y-10">
                <div>
                   <div className="flex items-center gap-3 mb-6">
                      <AlertTriangle className="text-accent" size={20} />
                      <h3 className="text-base font-black uppercase tracking-tight">Umbral de Existencias Críticas</h3>
                   </div>
                   
                   <p className="text-[11px] text-muted font-bold leading-relaxed uppercase mb-8 max-w-md opacity-70">
                     Define el porcentaje del stock mínimo a partir del cual un producto se marcará como <span className="text-red-500 underline underline-offset-4 decoration-2 tracking-widest font-black">CRÍTICO</span>. 
                     (Valor actual: {Math.round(criticalThreshold * 100)}%)
                   </p>

                   <div className="space-y-6">
                      <div className="flex justify-between items-end mb-2">
                        <span className="text-[10px] font-black uppercase tracking-[0.3em] text-muted">Ajuste de Sensibilidad</span>
                        <span className="text-4xl font-black text-accent">{Math.round(criticalThreshold * 100)}%</span>
                      </div>
                      
                      {accountType === 'COMPANY' ? (
                        <input 
                          type="range"
                          min="5"
                          max="80"
                          step="5"
                          value={criticalThreshold * 100}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) / 100;
                            setCriticalThreshold(val);
                            localStorage.setItem('sied_critical_threshold', val.toString());
                          }}
                          className="w-full h-2 bg-main/10 rounded-lg appearance-none cursor-pointer accent-accent"
                        />
                      ) : (
                        <div className="space-y-4">
                          <div className="w-full h-2 bg-main/5 rounded-lg overflow-hidden flex">
                            <div 
                              className="bg-accent/30 h-full transition-all" 
                              style={{ width: `${criticalThreshold * 100}%` }} 
                            />
                          </div>
                          <div className="flex items-center gap-3 bg-red-500/5 p-4 rounded-xl border border-red-500/10">
                            <Shield className="text-red-500" size={16} />
                            <p className="text-[9px] font-black text-red-500/80 uppercase tracking-widest leading-loose">
                              Solo el Administrador de la Empresa puede modificar el umbral global de existencias críticas.
                            </p>
                          </div>
                        </div>
                      )}
                      
                      <div className="flex justify-between text-[9px] font-black text-muted/40 uppercase tracking-[0.2em]">
                        <span>BAJA SENSIBILIDAD (5%)</span>
                        <span>ALTA SENSIBILIDAD (80%)</span>
                      </div>
                   </div>
                </div>

                <div className="pt-10 border-t border-main/10">
                   <div className="flex items-center gap-3 mb-4">
                      <Warehouse className="text-accent" size={20} />
                      <h3 className="text-base font-black uppercase tracking-tight">Base de Datos del Sistema (Cloud Firestore)</h3>
                   </div>
                   <p className="text-[11px] text-muted font-bold uppercase tracking-widest leading-relaxed mb-6">
                     Persistencia de datos en la nube: materias primas, insumos, productos terminados, recetas de producción, empaques, lotes y movimientos de Kardex.
                   </p>
                   
                   <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
                      <div className="bg-main/5 p-4 rounded-2xl border border-main/10">
                        <span className="text-[9px] font-black uppercase text-muted block mb-1">Productos</span>
                        <span className="text-xl font-black text-accent">{products.length}</span>
                      </div>
                      <div className="bg-main/5 p-4 rounded-2xl border border-main/10">
                        <span className="text-[9px] font-black uppercase text-muted block mb-1">Lotes Activos</span>
                        <span className="text-xl font-black text-main">{batches.length}</span>
                      </div>
                      <div className="bg-main/5 p-4 rounded-2xl border border-main/10">
                        <span className="text-[9px] font-black uppercase text-muted block mb-1">Transacciones</span>
                        <span className="text-xl font-black text-emerald-500">{transactions.length}</span>
                      </div>
                      <div className="bg-main/5 p-4 rounded-2xl border border-main/10">
                        <span className="text-[9px] font-black uppercase text-muted block mb-1">Órdenes Prod.</span>
                        <span className="text-xl font-black text-indigo-500">{productionOrders.length}</span>
                      </div>
                   </div>

                   <div className="bg-main/5 p-6 rounded-2xl border border-main/10 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                      <div>
                         <div className="text-[10px] font-black uppercase tracking-widest text-muted mb-1">Estado de Conexión</div>
                         <div className="flex items-center gap-2">
                            <div className="w-2.5 h-2.5 rounded-full bg-green-500 animate-pulse shadow-lg shadow-green-500/20" />
                            <span className="text-xs font-black text-main uppercase">
                              En Tiempo Real ({currentUser?.companyName || 'Empresa Activa'})
                            </span>
                         </div>
                      </div>

                      <button
                        onClick={handleSeedSystemData}
                        disabled={isSeedingData}
                        className="w-full sm:w-auto flex items-center justify-center gap-2 bg-accent text-white px-6 py-3 rounded-xl font-black uppercase text-[10px] tracking-widest hover:brightness-110 active:scale-95 transition-all shadow-lg shadow-brand-gold-shadow disabled:opacity-50"
                      >
                        <Sparkles size={14} className={isSeedingData ? 'animate-spin' : ''} />
                        <span>{isSeedingData ? 'Cargando Base de Datos...' : 'Cargar Datos Demo de Prueba'}</span>
                      </button>
                   </div>

                   {availableCompanies.length > 0 && (
                     <div className="bg-main/5 p-6 rounded-2xl border border-main/10 mt-6 space-y-4">
                       <div className="flex items-center justify-between">
                         <div className="flex items-center gap-2">
                           <Building2 size={16} className="text-accent" />
                           <span className="text-[10px] font-black uppercase tracking-widest text-muted">
                             Espacios de Trabajo en la Base de Datos ({availableCompanies.length})
                           </span>
                         </div>
                       </div>
                       <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                         {availableCompanies.map(comp => (
                           <button
                             key={comp.id}
                             onClick={() => handleSwitchCompany(comp.id, comp.name)}
                             className={cn(
                               "p-4 rounded-xl border text-left transition-all flex flex-col justify-between group",
                               comp.id === effectiveCompanyId
                                 ? "bg-accent/15 border-accent text-accent shadow-sm"
                                 : "bg-card border-main hover:border-accent/40 text-main"
                             )}
                           >
                             <div className="flex items-center justify-between mb-2">
                               <span className="font-black text-sm uppercase tracking-wide truncate">{comp.name}</span>
                               {comp.id === effectiveCompanyId && (
                                 <span className="text-[8px] font-black uppercase bg-accent text-white px-2 py-0.5 rounded-full">
                                   Activa
                                 </span>
                               )}
                             </div>
                             <div className="text-[10px] text-muted font-bold uppercase tracking-wider flex items-center gap-2">
                               <span>{comp.productCount} {comp.productCount === 1 ? 'producto' : 'productos'}</span>
                               <span>•</span>
                               <span>{comp.batchCount} lotes</span>
                             </div>
                           </button>
                         ))}
                       </div>
                     </div>
                   )}
                </div>

                <div className="pt-10 border-t border-main/10">
                   <div className="flex items-center gap-3 mb-4">
                      <Shield className="text-accent" size={20} />
                      <h3 className="text-base font-black uppercase tracking-tight">Seguridad de la Información</h3>
                   </div>
                   <p className="text-[11px] text-muted font-bold uppercase tracking-widest leading-relaxed mb-6">
                     Los cambios realizados en esta sección se guardan localmente para su dispositivo actual.
                   </p>
                   <div className="bg-main/5 p-6 rounded-2xl border border-main/10">
                      <div className="text-[10px] font-black uppercase tracking-widest text-muted mb-2">Estado del Servidor</div>
                      <div className="flex items-center gap-2">
                         <div className="w-2 h-2 rounded-full bg-green-500 shadow-lg shadow-green-500/20" />
                         <span className="text-xs font-black text-main uppercase">Operativo y Sincronizado</span>
                      </div>
                   </div>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </main>

      {/* Modals */}
      {showModal && (
        <div className="fixed inset-0 bg-main/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className={cn(
              "bg-card w-full rounded-3xl p-6 md:p-10 shadow-2xl border border-main overflow-y-auto max-h-[90vh]",
              showModal === 'product' ? "max-w-xl" : "max-w-md"
            )}
          >
            {showModal === 'product' ? (
              <>
                <h3 className="text-2xl font-black mb-6 uppercase tracking-tight text-accent">REGISTRAR NUEVO PRODUCTO</h3>
                <form onSubmit={handleAddProduct} className="space-y-4">
                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Nombre del Producto</label>
                    <input 
                      type="text" required
                      className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main"
                      onChange={e => setFormData({ ...formData, name: e.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Descripción (Opcional)</label>
                    <input 
                      type="text"
                      className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-sm"
                      value={formData.description}
                      onChange={e => setFormData({ ...formData, description: e.target.value })}
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Categoría</label>
                      <input 
                        type="text" required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs"
                        value={formData.category}
                        onChange={e => setFormData({ ...formData, category: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Stock Mínimo (Alerta)</label>
                      <input 
                        type="number" required min="0"
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs"
                        value={formData.minStock}
                        onFocus={(e) => e.target.select()}
                        onChange={e => setFormData({ ...formData, minStock: Number(e.target.value) })}
                      />
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Unidad de Medida</label>
                      <select 
                        required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs appearance-none"
                        value={formData.unit}
                        onChange={e => setFormData({ ...formData, unit: e.target.value as any })}
                      >
                        <option value="unidad" className="bg-sidebar">Unidad</option>
                        <option value="kilo" className="bg-sidebar">Kilo</option>
                        <option value="litro" className="bg-sidebar">Litro</option>
                      </select>
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Tipo de Inventario</label>
                    <div className="grid grid-cols-3 gap-2">
                        {['materia prima', 'producción', 'producto terminado'].map((t) => (
                            <button
                                key={t}
                                type="button"
                                onClick={() => setFormData({ ...formData, type: t as any })}
                                className={cn(
                                    "text-[9px] font-black uppercase py-2.5 rounded-lg border transition-all flex flex-col items-center justify-center",
                                    formData.type === t 
                                        ? "bg-accent text-white border-accent shadow-md shadow-brand-gold-shadow" 
                                        : "bg-main/5 text-muted border-main hover:bg-main/10"
                                )}
                            >
                                <span>{t}</span>
                                {formData.type === t && <span className="text-[8px] font-bold opacity-80">✓ Activo</span>}
                            </button>
                        ))}
                    </div>
                  </div>

                  {/* Consumo de Materiales / Receta (para producto terminado o en producción) */}
                  {(formData.type === 'producto terminado' || formData.type === 'producción') && (
                    <div className="p-4 rounded-2xl bg-amber-500/5 border border-amber-500/20 space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2 text-amber-500">
                          <Layers size={16} />
                          <span className="text-[10px] font-black uppercase tracking-wider">
                            Consumo de Materiales (Fórmula / Receta)
                          </span>
                        </div>
                        <span className="text-[9px] font-bold text-muted uppercase">
                          Por 1 {formData.unit}
                        </span>
                      </div>
                      <p className="text-[10px] text-muted font-medium">
                        Indique los insumos y materias primas que consume este producto para fabricarse:
                      </p>

                      {/* Lista de materiales agregados */}
                      {productRecipe.length > 0 && (
                        <div className="space-y-1.5">
                          {productRecipe.map((item, idx) => {
                            const mat = products.find(p => p.id === item.materialProductId);
                            return (
                              <div key={idx} className="flex items-center justify-between p-2 rounded-xl bg-sidebar border border-main text-xs">
                                <div className="flex items-center gap-2">
                                  <span className="w-5 h-5 rounded-full bg-accent/10 text-accent font-black text-[10px] flex items-center justify-center">
                                    {idx + 1}
                                  </span>
                                  <div>
                                    <span className="font-bold text-main">{mat?.name || item.materialName || 'Material'}</span>
                                    <span className="text-[9px] text-muted block uppercase tracking-wider">
                                      Consume: <b className="text-accent">{item.quantityNeeded}</b> {mat?.unit || item.unit || 'unidad'}s
                                    </span>
                                  </div>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setProductRecipe(prev => prev.filter((_, i) => i !== idx))}
                                  className="p-1 text-red-500 hover:bg-red-500/10 rounded-lg transition-colors"
                                  title="Quitar material"
                                >
                                  <X size={14} />
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* Selector para añadir material */}
                      <div className="flex flex-col sm:flex-row gap-2 pt-1">
                        <select
                          id="new-recipe-material-select"
                          className="flex-[2] bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main outline-none"
                          defaultValue=""
                        >
                          <option value="" disabled>Seleccione materia prima...</option>
                          {products
                            .filter(p => p.type === 'materia prima' || p.type === 'producción')
                            .map(p => (
                              <option key={p.id} value={p.id}>
                                {p.name} ({p.unit}s) - Stock: {productStock[p.id] || 0}
                              </option>
                            ))
                          }
                        </select>
                        <input
                          id="new-recipe-qty-input"
                          type="number"
                          step="0.01"
                          min="0.001"
                          placeholder="Cantidad"
                          className="flex-1 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main outline-none w-24"
                        />
                        <button
                          type="button"
                          onClick={() => {
                            const sel = document.getElementById('new-recipe-material-select') as HTMLSelectElement;
                            const qty = document.getElementById('new-recipe-qty-input') as HTMLInputElement;
                            if (!sel || !sel.value || !qty || Number(qty.value) <= 0) {
                              alert('Seleccione un material e ingrese la cantidad consumida.');
                              return;
                            }
                            const mat = products.find(p => p.id === sel.value);
                            setProductRecipe(prev => [
                              ...prev.filter(item => item.materialProductId !== sel.value),
                              {
                                materialProductId: sel.value,
                                materialName: mat?.name,
                                quantityNeeded: Number(qty.value),
                                unit: mat?.unit
                              }
                            ]);
                            qty.value = '';
                          }}
                          className="bg-accent text-white px-3 py-2 rounded-xl font-bold text-xs hover:brightness-110 flex items-center justify-center gap-1"
                        >
                          <Plus size={14} /> Añadir
                        </button>
                      </div>
                    </div>
                  )}

                  {/* Configuración de Empaque (Caja o Bulto) */}
                  {(formData.type === 'producto terminado' || formData.type === 'producción') && (
                    <div className="p-4 rounded-2xl bg-blue-500/5 border border-blue-500/20 space-y-3">
                      <div className="flex items-center gap-2 text-blue-500">
                        <Boxes size={16} />
                        <span className="text-[10px] font-black uppercase tracking-wider">
                          Empaque (Caja o Bulto y cuántas unidades lleva)
                        </span>
                      </div>
                      <div className="grid grid-cols-4 gap-1.5">
                        {[
                          { id: 'ninguno' as const, label: 'Sin Empaque' },
                          { id: 'caja' as const, label: 'En Cajas' },
                          { id: 'bulto' as const, label: 'En Bultos' },
                          { id: 'ambos' as const, label: 'Cajas y Bultos' }
                        ].map(opt => (
                          <button
                            key={opt.id}
                            type="button"
                            onClick={() => setProductPackaging(prev => ({ ...prev, packagingType: opt.id }))}
                            className={cn(
                              "text-[9px] font-black uppercase py-2 px-1 rounded-xl border transition-all text-center",
                              productPackaging.packagingType === opt.id
                                ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                                : "bg-sidebar text-muted border-main hover:bg-main/5"
                            )}
                          >
                            {opt.label}
                          </button>
                        ))}
                      </div>

                      {(productPackaging.packagingType === 'caja' || productPackaging.packagingType === 'ambos') && (
                        <div className="flex items-center justify-between gap-3 pt-1">
                          <label className="text-[10px] font-black text-muted uppercase tracking-wider flex items-center gap-1.5">
                            📦 ¿Cuántas unidades lleva cada <b>Caja</b>?
                          </label>
                          <input
                            type="number"
                            min="1"
                            placeholder="ej. 12"
                            value={productPackaging.unitsPerBox || ''}
                            onChange={e => setProductPackaging(prev => ({ ...prev, unitsPerBox: Number(e.target.value) || 0 }))}
                            className="w-28 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main focus:border-accent outline-none text-right"
                          />
                        </div>
                      )}

                      {(productPackaging.packagingType === 'bulto' || productPackaging.packagingType === 'ambos') && (
                        <div className="flex items-center justify-between gap-3 pt-1">
                          <label className="text-[10px] font-black text-muted uppercase tracking-wider flex items-center gap-1.5">
                            🧺 ¿Cuántas unidades lleva cada <b>Bulto</b>?
                          </label>
                          <input
                            type="number"
                            min="1"
                            placeholder="ej. 50"
                            value={productPackaging.unitsPerBundle || ''}
                            onChange={e => setProductPackaging(prev => ({ ...prev, unitsPerBundle: Number(e.target.value) || 0 }))}
                            className="w-28 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main focus:border-accent outline-none text-right"
                          />
                        </div>
                      )}
                    </div>
                  )}

                  <div className="flex gap-4 pt-4">
                    <button type="button" onClick={() => setShowModal(null)} className="flex-1 bg-main/5 text-muted py-4 rounded-xl font-bold border border-main uppercase text-xs">Cancelar</button>
                    <button type="submit" className="flex-[2] bg-accent text-white py-4 rounded-xl font-black tracking-widest hover:brightness-110 shadow-xl shadow-brand-gold-shadow text-xs">CREAR PRODUCTO</button>
                  </div>
                </form>
              </>
            ) : (
              <>
                <h3 className="text-2xl font-black mb-6 uppercase tracking-tight text-accent">
                  {showModal === 'entry' ? 'REGISTRAR ENTRADA' : 
                   showModal === 'exit' ? (exitItems.length > 0 ? 'REGISTRAR SALIDAS (MÚLTIPLE)' : 'REGISTRAR SALIDA') :
                   showModal === 'purchase' ? 'REGISTRAR COMPRA' : 
                   showModal === 'edit_batch' ? 'EDITAR LOTE / REGISTRO' :
                   (saleItems.length > 0 ? 'REGISTRAR VENTA MASIVA' : 'REGISTRAR VENTA')}
                </h3>
                <form 
                  onSubmit={
                    showModal === 'entry' ? handleAddEntry : 
                    showModal === 'exit' ? handleInitiateExit :
                    showModal === 'purchase' ? handleAddPurchase : 
                    showModal === 'edit_batch' ? handleEditBatch :
                    handleInitiateSale
                  } 
                  className="space-y-6"
                >
                  {(showModal === 'entry' || showModal === 'purchase' || showModal === 'exit' || showModal === 'sale' || showModal === 'edit_batch') && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2 flex items-center justify-between">
                        <span>Fecha del Movimiento</span>
                        <span className="text-[9px] text-accent font-bold uppercase tracking-wider">Sincroniza con Kardex</span>
                      </label>
                      <input 
                        type="date" required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                        value={formData.entryDate}
                        onChange={e => setFormData({ ...formData, entryDate: e.target.value })}
                      />
                    </div>
                  )}

                  {/* Purchase Multi-Item List */}
                  {showModal === 'purchase' && purchaseItems.length > 0 && (
                    <div className="space-y-3 mb-2 p-4 bg-sidebar/50 rounded-2xl border border-main border-dashed">
                      <div className="flex justify-between items-center">
                        <span className="text-[10px] font-black uppercase tracking-widest text-muted">Lista de Compra</span>
                        <span className="text-[9px] font-black text-accent bg-accent/10 px-2 py-0.5 rounded-full">{purchaseItems.length} items</span>
                      </div>
                      <div className="space-y-2 max-h-[120px] overflow-y-auto pr-2 custom-scrollbar">
                        {purchaseItems.map((item, idx) => (
                          <div key={idx} className="flex justify-between items-center p-3 bg-main/5 rounded-xl border border-main/10 group">
                            <div className="flex gap-3 items-center">
                              <div className="w-8 h-8 rounded-lg bg-accent/10 flex items-center justify-center text-accent font-black text-[10px]">
                                {idx + 1}
                              </div>
                              <div>
                                <div className="text-[10px] font-black uppercase text-main leading-tight">{item.productName}</div>
                                <div className="text-[9px] font-bold text-muted uppercase tracking-wider">
                                  {item.quantity} x Bs. {item.cost.toLocaleString()}
                                </div>
                              </div>
                            </div>
                            <button 
                              type="button" 
                              onClick={() => setPurchaseItems(prev => prev.filter((_, i) => i !== idx))}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-lg transition-all"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Sale Multi-Item List */}
                  {showModal === 'sale' && saleItems.length > 0 && (
                    <div className="space-y-3 mb-2 p-4 bg-sidebar/50 rounded-2xl border border-accent/30 border-dashed">
                      <div className="flex justify-between items-center">
                        <span className="text-[10px] font-black uppercase tracking-widest text-accent flex items-center gap-1.5">
                          <ShoppingCart size={13} />
                          Productos en la Venta
                        </span>
                        <span className="text-[9px] font-black text-accent bg-accent/10 px-2 py-0.5 rounded-full">{saleItems.length} items</span>
                      </div>
                      <div className="space-y-2 max-h-[130px] overflow-y-auto pr-2 custom-scrollbar">
                        {saleItems.map((item, idx) => (
                          <div key={idx} className="flex justify-between items-center p-3 bg-main/5 rounded-xl border border-main/10 group">
                            <div className="flex gap-3 items-center">
                              <div className="w-8 h-8 rounded-lg bg-accent/10 flex items-center justify-center text-accent font-black text-[10px]">
                                {idx + 1}
                              </div>
                              <div>
                                <div className="text-[10px] font-black uppercase text-main leading-tight">{item.productName}</div>
                                <div className="text-[9px] font-bold text-muted uppercase tracking-wider">
                                  {item.quantity} x Bs. {item.salePrice.toLocaleString()} = <span className="text-accent font-black">Bs. {(item.quantity * item.salePrice).toFixed(2)}</span>
                                </div>
                              </div>
                            </div>
                            <button 
                              type="button" 
                              onClick={() => setSaleItems(prev => prev.filter((_, i) => i !== idx))}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-lg transition-all"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Exit Multi-Item List */}
                  {showModal === 'exit' && exitItems.length > 0 && (
                    <div className="space-y-3 mb-2 p-4 bg-sidebar/50 rounded-2xl border border-amber-500/30 border-dashed">
                      <div className="flex justify-between items-center">
                        <span className="text-[10px] font-black uppercase tracking-widest text-amber-500 flex items-center gap-1.5">
                          <Boxes size={13} />
                          Productos en la Salida
                        </span>
                        <span className="text-[9px] font-black text-amber-500 bg-amber-500/10 px-2 py-0.5 rounded-full">{exitItems.length} items</span>
                      </div>
                      <div className="space-y-2 max-h-[130px] overflow-y-auto pr-2 custom-scrollbar">
                        {exitItems.map((item, idx) => (
                          <div key={idx} className="flex justify-between items-center p-3 bg-main/5 rounded-xl border border-main/10 group">
                            <div className="flex gap-3 items-center">
                              <div className="w-8 h-8 rounded-lg bg-amber-500/10 flex items-center justify-center text-amber-500 font-black text-[10px]">
                                {idx + 1}
                              </div>
                              <div>
                                <div className="text-[10px] font-black uppercase text-main leading-tight">{item.productName}</div>
                                <div className="text-[9px] font-bold text-amber-500 uppercase tracking-wider">
                                  {item.quantity} unidades
                                </div>
                              </div>
                            </div>
                            <button 
                              type="button" 
                              onClick={() => setExitItems(prev => prev.filter((_, i) => i !== idx))}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-lg transition-all"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {(showModal === 'purchase' || showModal === 'sale' || showModal === 'exit') && !selectedProductId && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Seleccionar Producto</label>
                      <select 
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                        value={selectedProductId || ''}
                        onChange={e => setSelectedProductId(e.target.value)}
                      >
                        <option value="">ELIJA UN PRODUCTO...</option>
                        {products.sort((a,b) => a.name.localeCompare(b.name)).map(p => (
                          <option key={p.id} value={p.id}>
                            {p.name} (Stock: {productStock[p.id] || 0} {p.unit}s)
                          </option>
                        ))}
                      </select>
                    </div>
                  )}

                  {selectedProductId && (
                    <div className="bg-main/5 p-4 rounded-2xl border border-main/10 flex items-center justify-between mb-2">
                        <div className="flex items-center gap-4">
                          <History size={20} className="text-accent" />
                          <div>
                              <div className="text-[10px] font-black text-muted uppercase tracking-widest leading-none mb-1">
                                {showModal === 'sale' ? 'Producto a vender:' : showModal === 'exit' ? 'Producto a descontar:' : 'Cargar a:'}
                              </div>
                              <div className="text-sm font-black text-main leading-none uppercase">{products.find(p => p.id === selectedProductId)?.name}</div>
                          </div>
                        </div>
                        {(showModal === 'purchase' || showModal === 'sale' || showModal === 'exit') && (
                          <button 
                            type="button"
                            onClick={() => setSelectedProductId(null)}
                            className="bg-main/10 text-muted p-2 rounded-lg hover:bg-main/20 text-[10px] font-black"
                          >
                            CAMBIAR
                          </button>
                        )}
                    </div>
                  )}

                  {showModal === 'edit_batch' && (
                    <div className="bg-orange-500/10 border border-orange-500/20 p-4 rounded-xl mb-4">
                      <div className="flex items-center gap-2 text-orange-500 mb-1">
                        <AlertTriangle size={14} />
                        <span className="text-[10px] font-black uppercase tracking-widest">Aviso de Integridad</span>
                      </div>
                      <p className="text-[10px] text-muted font-bold leading-relaxed uppercase">
                        La edición directa de un lote ajustará el saldo disponible y el costo histórico. 
                        Use esta función solo para correcciones administrativas.
                      </p>
                    </div>
                  )}

                  <div>
                    <div className="flex justify-between items-end mb-2">
                        <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em]">
                          {showModal === 'edit_batch' ? 'Cantidad Total del Lote' : 'Cantidad'}
                        </label>
                        {(showModal === 'sale' || showModal === 'exit') && selectedProductId && (
                            <span className={cn(
                                "text-[10px] font-black uppercase",
                                formData.quantity > (productStock[selectedProductId] || 0) ? "text-red-500" : "text-accent"
                            )}>
                                Disponible: {productStock[selectedProductId] || 0} {products.find(p => p.id === selectedProductId)?.unit}s
                            </span>
                        )}
                    </div>
                    <input 
                      type="number" 
                      required={(showModal === 'sale' ? saleItems.length === 0 : showModal === 'exit' ? exitItems.length === 0 : true)}
                      className={cn(
                        "w-full bg-black/5 dark:bg-black/30 border rounded-xl px-5 py-4 focus:outline-none font-bold text-main",
                        (showModal === 'sale' || showModal === 'exit') && selectedProductId && formData.quantity > (productStock[selectedProductId] || 0)
                          ? "border-red-500 focus:border-red-500" 
                          : "border-main focus:border-accent"
                      )}
                      value={formData.quantity}
                      onFocus={(e) => e.target.select()}
                      onKeyDown={(e) => {
                        if (['e', 'E', '+', '-'].includes(e.key)) e.preventDefault();
                      }}
                      onChange={e => setFormData({ ...formData, quantity: Number(e.target.value) })}
                    />
                    {(showModal === 'sale' || showModal === 'exit') && selectedProductId && formData.quantity > (productStock[selectedProductId] || 0) && (
                        <p className="text-[9px] font-black uppercase text-red-500 mt-2 flex items-center gap-1">
                            <span className="w-1 h-1 bg-red-500 rounded-full animate-pulse"></span>
                            Stock insuficiente para completar esta operación
                        </p>
                    )}
                  </div>

                  {(showModal !== 'exit' && showModal !== 'sale' && showModal !== 'edit_batch') && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Costo Unitario (Bs)</label>
                      <input 
                        type="number" step="0.01" required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main font-mono"
                        value={formData.cost}
                        onFocus={(e) => e.target.select()}
                        onKeyDown={(e) => {
                          if (['e', 'E', '+', '-'].includes(e.key)) e.preventDefault();
                        }}
                        onChange={e => setFormData({ ...formData, cost: Number(e.target.value) })}
                      />
                    </div>
                  )}

                  {showModal === 'edit_batch' && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Costo Unitario (Edición)</label>
                      <input 
                        type="number" step="0.01" required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main font-mono"
                        value={formData.cost}
                        onFocus={(e) => e.target.select()}
                        onKeyDown={(e) => {
                          if (['e', 'E', '+', '-'].includes(e.key)) e.preventDefault();
                        }}
                        onChange={e => setFormData({ ...formData, cost: Number(e.target.value) })}
                      />
                    </div>
                  )}

                  {showModal === 'sale' && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Precio de Venta Unitario (Bs)</label>
                      <input 
                        type="number" step="0.01" 
                        required={saleItems.length === 0 || (selectedProductId !== null && formData.quantity > 0)}
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main font-mono"
                        value={formData.cost} // Reusing cost state for price
                        onFocus={(e) => e.target.select()}
                        onKeyDown={(e) => {
                          if (['e', 'E', '+', '-'].includes(e.key)) e.preventDefault();
                        }}
                        onChange={e => setFormData({ ...formData, cost: Number(e.target.value) })}
                      />
                    </div>
                  )}

                  {(showModal === 'entry' || showModal === 'purchase' || showModal === 'edit_batch') && (
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Proveedor</label>
                      <input 
                        type="text" required
                        className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                        value={formData.supplier}
                        onChange={e => setFormData({ ...formData, supplier: e.target.value })}
                        placeholder="NOMBRE DEL PROVEEDOR"
                      />
                    </div>
                  )}

                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Descripción / Motivo</label>
                    <textarea 
                      required
                      className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main text-xs min-h-[80px]"
                      value={formData.productDescription}
                      onChange={e => setFormData({ ...formData, productDescription: e.target.value })}
                      placeholder="Indique brevemente el motivo o referencia del registro..."
                    />
                  </div>
                  
                  {showModal === 'purchase' && (
                    <div className="pt-2">
                      <button 
                        type="button"
                        onClick={addItemToPurchase}
                        disabled={!selectedProductId || formData.quantity <= 0}
                        className="w-full bg-sidebar border border-accent/20 text-accent py-3.5 rounded-xl font-black text-xs uppercase tracking-widest hover:bg-accent/5 transition-all disabled:opacity-50 flex items-center justify-center gap-2"
                      >
                        <Plus size={16} /> Añadir otro producto a esta compra
                      </button>
                    </div>
                  )}

                  {showModal === 'sale' && (
                    <div className="pt-2">
                      <button 
                        type="button"
                        onClick={addItemToSale}
                        disabled={!selectedProductId || formData.quantity <= 0 || (selectedProductId && formData.quantity > (productStock[selectedProductId] || 0))}
                        className="w-full bg-sidebar border border-accent/30 text-accent py-3.5 rounded-xl font-black text-xs uppercase tracking-widest hover:bg-accent/10 transition-all disabled:opacity-50 flex items-center justify-center gap-2"
                      >
                        <Plus size={16} /> Añadir otro producto a esta venta (Venta Masiva)
                      </button>
                    </div>
                  )}

                  {showModal === 'exit' && (
                    <div className="pt-2">
                      <button 
                        type="button"
                        onClick={addItemToExit}
                        disabled={!selectedProductId || formData.quantity <= 0 || (selectedProductId && formData.quantity > (productStock[selectedProductId] || 0))}
                        className="w-full bg-sidebar border border-amber-500/30 text-amber-500 py-3.5 rounded-xl font-black text-xs uppercase tracking-widest hover:bg-amber-500/10 transition-all disabled:opacity-50 flex items-center justify-center gap-2"
                      >
                        <Plus size={16} /> Añadir otro producto a esta salida (Salida Masiva)
                      </button>
                    </div>
                  )}
                  
                  <div className="flex gap-4 pt-4">
                    <button 
                      type="button"
                      onClick={() => {
                        setShowModal(null);
                        setPurchaseItems([]);
                        setSaleItems([]);
                        setExitItems([]);
                      }}
                      className="flex-1 bg-main/5 text-muted py-4 rounded-xl font-bold hover:bg-main/10 transition-all border border-main uppercase tracking-widest text-xs"
                    >
                      Cancelar
                    </button>
                    <button 
                      type="submit"
                      disabled={
                        (showModal === 'sale' && selectedProductId && formData.quantity > (productStock[selectedProductId] || 0)) ||
                        (showModal === 'exit' && selectedProductId && formData.quantity > (productStock[selectedProductId] || 0)) ||
                        (showModal === 'sale' && saleItems.length === 0 && (!selectedProductId || formData.quantity <= 0)) ||
                        (showModal === 'exit' && exitItems.length === 0 && (!selectedProductId || formData.quantity <= 0))
                      }
                      className={cn(
                        "flex-[2] py-4 rounded-xl font-black tracking-widest transition-all text-xs flex items-center justify-center gap-2",
                        ((showModal === 'sale' || showModal === 'exit') && selectedProductId && formData.quantity > (productStock[selectedProductId] || 0))
                          ? "bg-main/20 text-muted cursor-not-allowed border border-main/10"
                          : "bg-accent text-white hover:brightness-110 shadow-xl shadow-brand-gold-shadow"
                      )}
                    >
                      {showModal === 'purchase' 
                        ? (purchaseItems.length > 0 ? `TERMINAR COMPRA (${purchaseItems.length + (selectedProductId && formData.quantity > 0 ? 1 : 0)})` : 'CONFIRMAR COMPRA') 
                        : showModal === 'sale'
                        ? (saleItems.length > 0 ? `REVISAR IMPACTO (${saleItems.length + (selectedProductId && formData.quantity > 0 ? 1 : 0)} ITEMS)` : 'REVISAR IMPACTO EN STOCK')
                        : showModal === 'exit'
                        ? (exitItems.length > 0 ? `REVISAR IMPACTO (${exitItems.length + (selectedProductId && formData.quantity > 0 ? 1 : 0)} ITEMS)` : 'REVISAR IMPACTO EN STOCK')
                        : 'CONFIRMAR REGISTRO'}
                    </button>
                  </div>
                </form>
              </>
            )}
          </motion.div>
        </div>
      )}

      {/* Verification / Confirmation Modal */}
      <TransactionConfirmationModal
        isOpen={confirmationSimulation !== null}
        simulation={confirmationSimulation}
        movementDate={pendingTransactionData?.movementDate}
        onDateChange={(newDate) => {
          const nowTime = new Date().toTimeString().split(' ')[0];
          const resolved = newDate.includes('T') ? newDate : `${newDate}T${nowTime}`;
          setPendingTransactionData(prev => prev ? { ...prev, movementDate: resolved } : null);
        }}
        operatorName={currentUser?.name || currentUser?.email || 'Operador'}
        notes={pendingTransactionData?.notes}
        isSubmitting={isExecutingTransaction}
        onConfirm={handleConfirmAndExecuteTransaction}
        onCancel={handleCancelConfirmation}
        onBackToEdit={handleBackToEdit}
      />

      {/* Product History / Tarjeta Kardex Modal */}
      {showModal === 'product_history' && selectedProductId && (() => {
        const selectedProd = products.find(p => p.id === selectedProductId);
        const isFinished = selectedProd?.type === 'producto terminado';
        const prodBatches = batches.filter(b => b.productId === selectedProductId && b.remainingQuantity > 0);
        const currentStock = prodBatches.reduce((acc, b) => acc + b.remainingQuantity, 0);
        const currentValuation = prodBatches.reduce((acc, b) => acc + (b.remainingQuantity * b.unitCost), 0);
        const currentAvgCost = currentStock > 0 ? (currentValuation / currentStock) : 0;
        const movements = generateUnifiedKardex(products, batches, transactions)
          .filter(entry => entry.productId === selectedProductId);
        const packaging = selectedProd?.packaging;

        return (
          <div className="fixed inset-0 bg-main/80 backdrop-blur-md z-50 flex items-center justify-center p-3 md:p-6">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              className={cn(
                "bg-card w-full rounded-3xl p-6 md:p-8 shadow-2xl border border-main overflow-hidden flex flex-col max-h-[92vh]",
                isFinished ? "max-w-5xl" : "max-w-4xl"
              )}
            >
              {/* Header */}
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6 pb-4 border-b border-main/10">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-xl md:text-2xl font-black uppercase tracking-tight text-accent flex items-center gap-2">
                      <ClipboardList size={22} />
                      {isFinished ? 'TARJETA KARDEX VALORADA (COSTO PROMEDIO EN SALIDA)' : 'HISTORIAL DE MOVIMIENTOS'}
                    </h3>
                    {isFinished && (
                      <span className="bg-emerald-500/10 text-emerald-500 border border-emerald-500/30 text-[9px] font-black uppercase px-2.5 py-0.5 rounded-full tracking-wider flex items-center gap-1 shadow-sm">
                        <PackageCheck size={12} /> Producto Terminado
                      </span>
                    )}
                  </div>
                  
                  <div className="flex flex-wrap items-center gap-2 md:gap-3 mt-2 text-xs">
                    <span className="font-black text-main uppercase text-sm">{selectedProd?.name || 'Producto'}</span>
                    <span className="text-muted font-bold">({selectedProd?.unit || 'unidad'}s)</span>
                    <span className="opacity-20 text-main">•</span>
                    <span className="text-muted font-bold uppercase">{selectedProd?.category || 'General'}</span>
                    
                    {packaging && packaging.packagingType !== 'ninguno' && (
                      <>
                        <span className="opacity-20 text-main">•</span>
                        {packaging.unitsPerBox && packaging.unitsPerBox > 0 && (
                          <span className="bg-blue-500/10 text-blue-600 dark:text-blue-400 px-2 py-0.5 rounded font-black text-[10px] flex items-center gap-1 border border-blue-500/20">
                            📦 {packaging.unitsPerBox} un/caja
                          </span>
                        )}
                        {packaging.unitsPerBundle && packaging.unitsPerBundle > 0 && (
                          <span className="bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 px-2 py-0.5 rounded font-black text-[10px] flex items-center gap-1 border border-indigo-500/20">
                            🧺 {packaging.unitsPerBundle} un/bulto
                          </span>
                        )}
                      </>
                    )}

                    {selectedProd?.recipe && selectedProd.recipe.length > 0 && (
                      <>
                        <span className="opacity-20 text-main">•</span>
                        <span className="bg-amber-500/10 text-amber-600 dark:text-amber-400 px-2 py-0.5 rounded font-black text-[10px] flex items-center gap-1 border border-amber-500/20">
                          <Layers size={10} /> {selectedProd.recipe.length} insumos de consumo
                        </span>
                      </>
                    )}
                  </div>
                </div>

                <button 
                  onClick={() => setShowModal(null)} 
                  className="p-2 hover:bg-main/10 rounded-xl text-muted hover:text-main transition-colors ml-auto sm:ml-0"
                  title="Cerrar"
                >
                  <X size={22} />
                </button>
              </div>

              {/* Informational banner for raw materials */}
              {!isFinished && (
                <div className="mb-4 p-3 rounded-xl bg-blue-500/10 border border-blue-500/20 text-blue-600 dark:text-blue-400 text-xs font-medium flex items-center gap-2">
                  <span>ℹ️</span>
                  <span>
                    La <b>Tarjeta Kardex Contable Valuada</b> oficial está enfocada en <b>Productos Terminados</b> valorados al costo promedio de salida. A continuación se detalla el control físico de entradas y consumos de esta materia prima o insumo.
                  </span>
                </div>
              )}

              {/* KPI Summary Cards for Finished Product */}
              {isFinished && (
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
                  <div className="p-3.5 rounded-2xl bg-sidebar border border-main">
                    <span className="text-[9px] font-black uppercase text-muted tracking-wider block mb-1">
                      Stock en Almacén
                    </span>
                    <span className="text-base md:text-lg font-black text-main">
                      {currentStock.toLocaleString()} {selectedProd?.unit}s
                    </span>
                  </div>

                  <div className="p-3.5 rounded-2xl bg-sidebar border border-main">
                    <span className="text-[9px] font-black uppercase text-muted tracking-wider block mb-1">
                      Equivalencia de Empaque
                    </span>
                    <span className="text-xs md:text-sm font-black text-blue-500 block truncate">
                      {packaging?.unitsPerBox && packaging.unitsPerBox > 0 
                        ? `📦 ${(currentStock / packaging.unitsPerBox).toFixed(1)} Cajas (${packaging.unitsPerBox} c/u)`
                        : packaging?.unitsPerBundle && packaging.unitsPerBundle > 0
                        ? `🧺 ${(currentStock / packaging.unitsPerBundle).toFixed(1)} Bultos (${packaging.unitsPerBundle} c/u)`
                        : 'Sin empaque'}
                    </span>
                  </div>

                  <div className="p-3.5 rounded-2xl bg-sidebar border border-main">
                    <span className="text-[9px] font-black uppercase text-emerald-500 tracking-wider block mb-1">
                      Valor Total en Almacén
                    </span>
                    <span className="text-base md:text-lg font-black text-emerald-500">
                      Bs. {currentValuation.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </span>
                  </div>

                  <div className="p-3.5 rounded-2xl bg-sidebar border border-main">
                    <span className="text-[9px] font-black uppercase text-accent tracking-wider block mb-1">
                      Costo Promedio Ponderado
                    </span>
                    <span className="text-base md:text-lg font-black text-accent">
                      Bs. {currentAvgCost.toFixed(2)}
                    </span>
                  </div>
                </div>
              )}

              {/* Kardex Movements Table */}
              <div className="flex-1 overflow-auto bg-black/5 dark:bg-black/20 rounded-2xl border border-main/10 custom-scrollbar">
                <table className="w-full text-left text-xs min-w-[750px]">
                  <thead className="text-[10px] font-black uppercase tracking-widest text-muted bg-main/5 sticky top-0 z-10 border-b border-main/10 backdrop-blur-sm">
                    {isFinished ? (
                      <>
                        <tr>
                          <th className="px-4 py-3" rowSpan={2}>Fecha/Hora</th>
                          <th className="px-4 py-3" rowSpan={2}>Referencia / Motivo</th>
                          <th className="px-4 py-3 text-center border-l border-main/10 bg-black/5" colSpan={3}>Entradas (Producción)</th>
                          <th className="px-4 py-3 text-center border-l border-main/10" colSpan={3}>Salidas (Costo Promedio)</th>
                          <th className="px-4 py-3 text-center border-l border-main/10 bg-black/5" colSpan={3}>Saldos Valorados (Costo Promedio)</th>
                          <th className="px-4 py-3 text-right border-l border-main/10" rowSpan={2}>Responsable</th>
                        </tr>
                        <tr className="border-t border-main/10 bg-main/5">
                          <th className="px-3 py-2 border-l border-main/10 text-center opacity-70">Cant.</th>
                          <th className="px-3 py-2 text-center opacity-70">Cost.U</th>
                          <th className="px-3 py-2 text-center opacity-70">Total</th>
                          <th className="px-3 py-2 border-l border-main/10 text-center opacity-70">Cant.</th>
                          <th className="px-3 py-2 text-center opacity-70" title="Costo promedio ponderado al momento de salida">C. Prom.</th>
                          <th className="px-3 py-2 text-center opacity-70">Total</th>
                          <th className="px-3 py-2 border-l border-main/10 text-center opacity-70">Cant.</th>
                          <th className="px-3 py-2 text-center opacity-70">Prom.P</th>
                          <th className="px-3 py-2 text-center opacity-70 text-accent">Valor Bs.</th>
                        </tr>
                      </>
                    ) : (
                      <tr>
                        <th className="px-6 py-4">Fecha</th>
                        <th className="px-6 py-4">Referencia</th>
                        <th className="px-6 py-4 text-center">Entrada</th>
                        <th className="px-6 py-4 text-center">Salida</th>
                        <th className="px-6 py-4 text-center text-accent">P. Venta</th>
                        <th className="px-6 py-4 text-center">Saldo</th>
                        <th className="px-6 py-4 text-right">Responsable</th>
                      </tr>
                    )}
                  </thead>
                  <tbody className="divide-y divide-main/5 font-bold">
                    {movements.length > 0 ? (
                      movements.map((entry, idx) => (
                        <tr key={`${entry.date}-${idx}`} className="hover:bg-accent/5 transition-colors">
                          <td className="px-4 py-3 whitespace-nowrap opacity-70 text-[11px]">
                            {format(new Date(entry.date), 'dd/MM/yy HH:mm')}
                          </td>
                          <td className="px-4 py-3">
                            <div className="truncate max-w-[220px]" title={entry.reference}>
                              {entry.reference}
                            </div>
                          </td>

                          {isFinished ? (
                            <>
                              {/* Entradas */}
                              <td className="px-3 py-3 border-l border-main/10 text-green-500 font-black text-center">
                                {entry.entryQty > 0 ? (
                                  <div>
                                    <span>+{entry.entryQty}</span>
                                    {packaging?.unitsPerBox && packaging.unitsPerBox > 0 && (
                                      <span className="text-[8px] font-normal text-muted block">
                                        ({(entry.entryQty / packaging.unitsPerBox).toFixed(1)} cj)
                                      </span>
                                    )}
                                  </div>
                                ) : '-'}
                              </td>
                              <td className="px-3 py-3 text-main text-center font-mono text-[11px] opacity-80">
                                {entry.entryUnitCost > 0 ? `Bs. ${entry.entryUnitCost.toFixed(2)}` : ''}
                              </td>
                              <td className="px-3 py-3 text-green-500 text-center font-black">
                                {entry.entryTotal > 0 ? `Bs. ${entry.entryTotal.toFixed(2)}` : ''}
                              </td>

                              {/* Salidas */}
                              <td className="px-3 py-3 border-l border-main/10 text-red-500 font-black text-center">
                                {entry.exitQty > 0 ? (
                                  <div>
                                    <span>-{entry.exitQty}</span>
                                    {packaging?.unitsPerBox && packaging.unitsPerBox > 0 && (
                                      <span className="text-[8px] font-normal text-muted block">
                                        ({(entry.exitQty / packaging.unitsPerBox).toFixed(1)} cj)
                                      </span>
                                    )}
                                  </div>
                                ) : '-'}
                              </td>
                              <td className="px-3 py-3 text-main text-center font-mono text-[11px] opacity-80">
                                {entry.exitUnitCost > 0 ? `Bs. ${entry.exitUnitCost.toFixed(2)}` : ''}
                              </td>
                              <td className="px-3 py-3 text-red-500 text-center font-black">
                                {entry.exitTotal > 0 ? `Bs. ${entry.exitTotal.toFixed(2)}` : ''}
                              </td>

                              {/* Saldos */}
                              <td className="px-3 py-3 border-l border-main/10 font-black text-accent text-center">
                                <div>
                                  <span>{entry.balanceQty}</span>
                                  {packaging?.unitsPerBox && packaging.unitsPerBox > 0 && (
                                    <span className="text-[8px] font-normal text-muted block">
                                      ({(entry.balanceQty / packaging.unitsPerBox).toFixed(1)} cj)
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="px-3 py-3 text-muted text-center font-mono text-[11px] opacity-60">
                                Bs. {entry.balanceUnitCost.toFixed(2)}
                              </td>
                              <td className="px-3 py-3 text-accent font-black text-center">
                                Bs. {entry.balanceTotal.toFixed(2)}
                              </td>
                            </>
                          ) : (
                            <>
                              <td className="px-6 py-4 text-center text-green-500">
                                {entry.entryQty > 0 ? `+${entry.entryQty}` : '-'}
                              </td>
                              <td className="px-6 py-4 text-center text-red-500">
                                {entry.exitQty > 0 ? `-${entry.exitQty}` : '-'}
                              </td>
                              <td className="px-6 py-4 text-center text-accent font-black">
                                {entry.salePrice ? `Bs. ${entry.salePrice.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-6 py-4 text-center text-accent">
                                {entry.balanceQty}
                              </td>
                            </>
                          )}

                          <td className="px-4 py-3 text-right opacity-60 italic text-[10px] whitespace-nowrap">
                            {entry.performedBy || 'Sistema'}
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={isFinished ? 11 : 7} className="p-10 text-center text-muted font-bold uppercase tracking-widest opacity-40">
                          Sin movimientos registrados para este producto
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              
              {/* Footer Actions */}
              <div className="mt-5 pt-4 border-t border-main/10 flex flex-col sm:flex-row justify-between items-center gap-3">
                <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                  <button 
                    type="button"
                    onClick={() => handleExportSingleProductKardex(selectedProductId)}
                    className="w-full sm:w-auto flex items-center justify-center gap-2 px-5 py-2.5 bg-sidebar border border-accent/40 text-accent rounded-xl font-black uppercase tracking-widest text-[10px] md:text-xs hover:bg-accent hover:text-white transition-all shadow-sm active:scale-95"
                  >
                    <FileSpreadsheet size={16} />
                    Exportar Tarjeta Kardex (Excel)
                  </button>

                  {isFinished && (
                    <button
                      type="button"
                      onClick={() => {
                        setKardexFilterProduct(selectedProductId);
                        setActiveTab('kardex');
                        setShowModal(null);
                      }}
                      className="w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-2.5 bg-emerald-500/10 border border-emerald-500/30 text-emerald-600 dark:text-emerald-400 rounded-xl font-black uppercase tracking-widest text-[10px] md:text-xs hover:bg-emerald-500 hover:text-white transition-all shadow-sm active:scale-95"
                    >
                      <ClipboardList size={16} />
                      Ver en Tarjeta Kardex General
                    </button>
                  )}
                </div>

                <button 
                  onClick={() => setShowModal(null)}
                  className="w-full sm:w-auto px-7 py-2.5 bg-accent text-white rounded-xl font-black uppercase tracking-widest text-xs hover:brightness-110 shadow-lg shadow-brand-gold-shadow"
                >
                  Cerrar
                </button>
              </div>
            </motion.div>
          </div>
        );
      })()}

      {/* People Modal */}
      {showPeopleModal && (
        <div className="fixed inset-0 bg-main/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white w-full max-w-lg rounded-[2.5rem] p-10 shadow-2xl border border-main/5 overflow-hidden"
          >
             <h3 className="text-2xl font-black mb-8 uppercase tracking-tighter text-accent">
               {newPerson.id ? 'EDITAR COLABORADOR' : 'NUEVO COLABORADOR'}
             </h3>
             <form onSubmit={handlePeopleSubmit} className="space-y-6">
                <div className="space-y-4">
                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Nombre del Colaborador</label>
                    <input 
                      type="text" required
                      className="w-full bg-sidebar border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main"
                      placeholder="ej. Eduardo Torrealba"
                      value={newPerson.name || ''}
                      onChange={e => setNewPerson({ ...newPerson, name: e.target.value })}
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Correo Electrónico</label>
                    <input 
                      type="email" required
                      className="w-full bg-sidebar border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main"
                      placeholder="ej. colaborador@empresa.com"
                      value={newPerson.email || ''}
                      onChange={e => setNewPerson({ ...newPerson, email: e.target.value })}
                    />
                  </div>
                  
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Rol Asignado</label>
                      <select 
                        className="w-full bg-sidebar border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main appearance-none"
                        value={newPerson.role}
                        onChange={e => setNewPerson({ ...newPerson, role: e.target.value as UserRole })}
                      >
                        <option value="OPERATOR">OPERADOR</option>
                        <option value="ADMIN">ADMINISTRADOR</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">Estado</label>
                      <select 
                        className="w-full bg-sidebar border border-main rounded-xl px-5 py-4 focus:outline-none focus:border-accent font-bold text-main appearance-none"
                        value={newPerson.status}
                        onChange={e => setNewPerson({ ...newPerson, status: e.target.value as any })}
                      >
                        <option value="active">ACTIVO</option>
                        <option value="inactive">INACTIVO</option>
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-4">Permisos por Área</label>
                    <div className="grid grid-cols-2 gap-3">
                        {(['INVENTARIO', 'PRODUCCION', 'REPORTES', 'CONFIGURACION'] as AreaPermission[]).map((area) => (
                          <button
                            key={area}
                            type="button"
                            onClick={() => toggleArea(area)}
                            className={cn(
                              "flex items-center justify-between px-4 py-3 rounded-xl border font-bold text-[10px] uppercase tracking-wider transition-all",
                              newPerson.areas?.includes(area)
                                ? "bg-accent text-white border-accent shadow-lg shadow-brand-gold-shadow"
                                : "bg-sidebar text-muted border-main hover:border-accent/40"
                            )}
                          >
                            {area}
                            {newPerson.areas?.includes(area) && <CheckCircle size={14} />}
                          </button>
                        ))}
                    </div>
                  </div>
                </div>

                <div className="flex gap-4 pt-6">
                  <button 
                    type="button" 
                    onClick={() => { setShowPeopleModal(false); setNewPerson({ areas: [], role: 'OPERATOR' }); }} 
                    className="flex-1 bg-main/5 text-muted py-4 rounded-xl font-bold border border-main uppercase tracking-widest text-xs"
                  >
                    Cancelar
                  </button>
                  <button 
                    type="submit" 
                    className="flex-[2] bg-accent text-white py-4 rounded-xl font-black tracking-widest hover:brightness-110 shadow-xl shadow-brand-gold-shadow uppercase text-xs"
                  >
                    {newPerson.id ? 'Guardar Cambios' : 'Añadir al Equipo'}
                  </button>
                </div>
             </form>
          </motion.div>
        </div>
      )}

      {/* Modal: Fabricar / Producir Producto Terminado */}
      {showModal === 'produce' && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-xl rounded-3xl p-6 md:p-8 border border-main shadow-2xl overflow-y-auto max-h-[92vh]"
          >
            {/* Modal Header */}
            <div className="flex items-center justify-between pb-4 border-b border-main/20 mb-6">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-500">
                  <Factory size={22} />
                </div>
                <div>
                  <h3 className="text-xl font-black uppercase tracking-tight text-main">
                    Fabricar Producto Terminado
                  </h3>
                  <p className="text-muted text-[11px] font-bold">
                    Consume materias primas por PEPS y genera stock en unidades, cajas o bultos
                  </p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setShowModal(null)}
                className="p-2 text-muted hover:text-main hover:bg-main/5 rounded-xl transition-all"
                title="Cerrar"
              >
                <X size={20} />
              </button>
            </div>

            <form onSubmit={handleExecuteProductionOrder} className="space-y-5">
              {/* Selector de Producto Terminado */}
              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Producto Terminado a Producir *
                </label>
                <select
                  required
                  value={produceForm.productId}
                  onChange={e => {
                    const chosen = products.find(p => p.id === e.target.value);
                    const defaultMode = (chosen?.packaging && chosen.packaging.packagingType === 'caja') 
                      ? 'cajas' 
                      : (chosen?.packaging && chosen.packaging.packagingType === 'bulto')
                      ? 'bultos'
                      : 'unidades';
                    setProduceForm({
                      ...produceForm,
                      productId: e.target.value,
                      producedIn: defaultMode,
                      boxCount: 1,
                      bundleCount: 1,
                      quantityToProduce: 1
                    });
                  }}
                  className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main focus:outline-none focus:border-accent text-sm"
                >
                  <option value="" disabled>Seleccione producto terminado...</option>
                  {products
                    .filter(p => p.type === 'producto terminado' || p.type === 'producción')
                    .map(p => (
                      <option key={p.id} value={p.id}>
                        {p.name} ({p.unit}s) - Stock actual: {productStock[p.id] || 0}
                      </option>
                    ))
                  }
                </select>
              </div>

              {/* Detalle del producto seleccionado y su configuración */}
              {(() => {
                const selectedProd = products.find(p => p.id === produceForm.productId);
                if (!selectedProd) return null;

                const recipe = selectedProd.recipe || [];
                const packaging = selectedProd.packaging;

                // Calculate total base units according to mode
                let effectiveUnits = Number(produceForm.quantityToProduce) || 0;
                if (produceForm.producedIn === 'cajas' && packaging?.unitsPerBox) {
                  effectiveUnits = (Number(produceForm.boxCount) || 0) * packaging.unitsPerBox;
                } else if (produceForm.producedIn === 'bultos' && packaging?.unitsPerBundle) {
                  effectiveUnits = (Number(produceForm.bundleCount) || 0) * packaging.unitsPerBundle;
                }

                // Check material feasibility
                const materialsStatus = recipe.map(item => {
                  const needed = item.quantityNeeded * effectiveUnits;
                  const available = productStock[item.materialProductId] || 0;
                  const mat = products.find(p => p.id === item.materialProductId);
                  return {
                    ...item,
                    matName: mat?.name || item.materialName || 'Material',
                    unit: mat?.unit || item.unit || 'unidad',
                    needed,
                    available,
                    isSufficient: available >= needed
                  };
                });

                const allAvailable = materialsStatus.every(m => m.isSufficient);

                return (
                  <div className="space-y-4">
                    {/* Modo de Medición: Unidades, Cajas o Bultos */}
                    <div className="p-4 rounded-2xl bg-sidebar border border-main space-y-3">
                      <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em]">
                        ¿Cómo deseas registrar la producción?
                      </label>
                      <div className="grid grid-cols-3 gap-2">
                        <button
                          type="button"
                          onClick={() => setProduceForm({ ...produceForm, producedIn: 'unidades' })}
                          className={cn(
                            "py-2.5 px-2 rounded-xl border text-[10px] font-black uppercase transition-all flex flex-col items-center justify-center gap-0.5",
                            produceForm.producedIn === 'unidades'
                              ? "bg-accent text-white border-accent shadow-md shadow-brand-gold-shadow"
                              : "bg-main/5 text-muted border-main hover:bg-main/10"
                          )}
                        >
                          <span>Por {selectedProd.unit}s</span>
                          <span className="text-[8px] opacity-70">Unidades sueltas</span>
                        </button>
                        
                        <button
                          type="button"
                          disabled={!packaging || (packaging.packagingType !== 'caja' && packaging.packagingType !== 'ambos')}
                          onClick={() => setProduceForm({ ...produceForm, producedIn: 'cajas' })}
                          className={cn(
                            "py-2.5 px-2 rounded-xl border text-[10px] font-black uppercase transition-all flex flex-col items-center justify-center gap-0.5",
                            produceForm.producedIn === 'cajas'
                              ? "bg-blue-600 text-white border-blue-600 shadow-md"
                              : "bg-main/5 text-muted border-main hover:bg-main/10",
                            (!packaging || (packaging.packagingType !== 'caja' && packaging.packagingType !== 'ambos')) && "opacity-40 cursor-not-allowed"
                          )}
                        >
                          <span>Por Cajas</span>
                          <span className="text-[8px] opacity-80">
                            {packaging?.unitsPerBox ? `${packaging.unitsPerBox} ${selectedProd.unit}s c/u` : 'Sin config'}
                          </span>
                        </button>

                        <button
                          type="button"
                          disabled={!packaging || (packaging.packagingType !== 'bulto' && packaging.packagingType !== 'ambos')}
                          onClick={() => setProduceForm({ ...produceForm, producedIn: 'bultos' })}
                          className={cn(
                            "py-2.5 px-2 rounded-xl border text-[10px] font-black uppercase transition-all flex flex-col items-center justify-center gap-0.5",
                            produceForm.producedIn === 'bultos'
                              ? "bg-indigo-600 text-white border-indigo-600 shadow-md"
                              : "bg-main/5 text-muted border-main hover:bg-main/10",
                            (!packaging || (packaging.packagingType !== 'bulto' && packaging.packagingType !== 'ambos')) && "opacity-40 cursor-not-allowed"
                          )}
                        >
                          <span>Por Bultos</span>
                          <span className="text-[8px] opacity-80">
                            {packaging?.unitsPerBundle ? `${packaging.unitsPerBundle} ${selectedProd.unit}s c/u` : 'Sin config'}
                          </span>
                        </button>
                      </div>

                      {/* Inputs de Cantidad según el Modo */}
                      {produceForm.producedIn === 'unidades' && (
                        <div className="pt-2">
                          <label className="block text-[10px] font-black text-muted uppercase tracking-wider mb-1.5">
                            Cantidad de {selectedProd.unit}s a fabricar:
                          </label>
                          <input
                            type="number"
                            min="1"
                            step="1"
                            required
                            value={produceForm.quantityToProduce}
                            onFocus={e => e.target.select()}
                            onChange={e => setProduceForm({ ...produceForm, quantityToProduce: Number(e.target.value) || 0 })}
                            className="w-full bg-main/5 border border-main rounded-xl px-4 py-3 font-bold text-main focus:border-accent outline-none text-base font-mono"
                          />
                        </div>
                      )}

                      {produceForm.producedIn === 'cajas' && (
                        <div className="pt-2 space-y-2">
                          <div className="flex items-center justify-between">
                            <label className="text-[10px] font-black text-muted uppercase tracking-wider">
                              📦 Número de <b>Cajas</b> a fabricar:
                            </label>
                            <span className="text-[10px] font-mono text-blue-500 font-bold">
                              Cada caja lleva {packaging?.unitsPerBox || 1} {selectedProd.unit}s
                            </span>
                          </div>
                          <input
                            type="number"
                            min="1"
                            step="1"
                            required
                            value={produceForm.boxCount}
                            onFocus={e => e.target.select()}
                            onChange={e => setProduceForm({ ...produceForm, boxCount: Number(e.target.value) || 0 })}
                            className="w-full bg-main/5 border border-main rounded-xl px-4 py-3 font-bold text-main focus:border-blue-500 outline-none text-base font-mono"
                          />
                          <div className="p-2.5 rounded-lg bg-blue-500/10 text-blue-600 dark:text-blue-400 font-bold text-xs flex justify-between">
                            <span>Total de {selectedProd.unit}s producidas:</span>
                            <span className="font-mono text-sm">{effectiveUnits} {selectedProd.unit}s</span>
                          </div>
                        </div>
                      )}

                      {produceForm.producedIn === 'bultos' && (
                        <div className="pt-2 space-y-2">
                          <div className="flex items-center justify-between">
                            <label className="text-[10px] font-black text-muted uppercase tracking-wider">
                              🧺 Número de <b>Bultos</b> a fabricar:
                            </label>
                            <span className="text-[10px] font-mono text-indigo-500 font-bold">
                              Cada bulto lleva {packaging?.unitsPerBundle || 1} {selectedProd.unit}s
                            </span>
                          </div>
                          <input
                            type="number"
                            min="1"
                            step="1"
                            required
                            value={produceForm.bundleCount}
                            onFocus={e => e.target.select()}
                            onChange={e => setProduceForm({ ...produceForm, bundleCount: Number(e.target.value) || 0 })}
                            className="w-full bg-main/5 border border-main rounded-xl px-4 py-3 font-bold text-main focus:border-indigo-500 outline-none text-base font-mono"
                          />
                          <div className="p-2.5 rounded-lg bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 font-bold text-xs flex justify-between">
                            <span>Total de {selectedProd.unit}s producidas:</span>
                            <span className="font-mono text-sm">{effectiveUnits} {selectedProd.unit}s</span>
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Espacio de Consumo de Materiales en Tiempo Real */}
                    <div className="p-4 rounded-2xl bg-amber-500/5 border border-amber-500/20 space-y-3">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2 text-amber-500">
                          <Layers size={16} />
                          <span className="text-[10px] font-black uppercase tracking-wider">
                            Materiales Consumidos en esta Orden ({effectiveUnits} {selectedProd.unit}s)
                          </span>
                        </div>
                        {recipe.length === 0 && (
                          <span className="text-[10px] font-bold text-red-500">Sin receta configurada</span>
                        )}
                      </div>

                      {recipe.length === 0 ? (
                        <div className="p-3 rounded-xl bg-red-500/10 border border-red-500/20 text-red-500 text-xs">
                          <p className="font-bold">Este producto terminado aún no tiene materiales asignados.</p>
                          <p className="text-[11px] mt-1 text-muted">
                            Haz clic en el botón «Editar» del producto para agregar los materiales que consume por unidad.
                          </p>
                        </div>
                      ) : (
                        <div className="space-y-2">
                          {materialsStatus.map((mat, idx) => (
                            <div 
                              key={idx} 
                              className={cn(
                                "p-3 rounded-xl border flex items-center justify-between text-xs transition-all",
                                mat.isSufficient 
                                  ? "bg-sidebar border-main" 
                                  : "bg-red-500/10 border-red-500/30 text-red-500"
                              )}
                            >
                              <div>
                                <span className="font-bold text-main block">{mat.matName}</span>
                                <span className="text-[10px] text-muted font-mono">
                                  Consume {mat.quantityNeeded} × {effectiveUnits} = <b className="text-accent">{mat.needed.toFixed(2)}</b> {mat.unit}s
                                </span>
                              </div>
                              <div className="text-right">
                                <span className={cn(
                                  "text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider",
                                  mat.isSufficient ? "bg-emerald-500/10 text-emerald-500" : "bg-red-500 text-white font-black"
                                )}>
                                  {mat.isSufficient ? `Stock OK (${mat.available})` : `Faltan ${(mat.needed - mat.available).toFixed(2)}`}
                                </span>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}

                      {!allAvailable && recipe.length > 0 && (
                        <p className="text-[11px] font-bold text-red-500 flex items-center gap-1.5 pt-1">
                          <AlertTriangle size={14} /> No cuentas con suficiente stock de materia prima para esta cantidad.
                        </p>
                      )}
                    </div>

                    {/* Fecha y Notas */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                          Fecha de Producción
                        </label>
                        <input
                          type="date"
                          required
                          value={produceForm.productionDate}
                          onChange={e => setProduceForm({ ...produceForm, productionDate: e.target.value })}
                          className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main text-xs uppercase"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                          Observaciones / Lote Fabr.
                        </label>
                        <input
                          type="text"
                          placeholder="ej. Turno mañana, lote piloto..."
                          value={produceForm.notes}
                          onChange={e => setProduceForm({ ...produceForm, notes: e.target.value })}
                          className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main text-xs"
                        />
                      </div>
                    </div>

                    {/* Acciones del Modal */}
                    <div className="flex gap-4 pt-4 border-t border-main/20">
                      <button 
                        type="button" 
                        onClick={() => setShowModal(null)}
                        className="flex-1 bg-main/5 text-muted py-3.5 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest"
                      >
                        Cancelar
                      </button>
                      <button 
                        type="submit"
                        disabled={effectiveUnits <= 0 || recipe.length === 0 || !allAvailable}
                        className="flex-[2] bg-emerald-600 hover:bg-emerald-500 text-white py-3.5 rounded-xl font-black tracking-widest hover:brightness-110 shadow-lg shadow-emerald-950/20 uppercase text-xs disabled:opacity-40 disabled:cursor-not-allowed transition-all flex items-center justify-center gap-2"
                      >
                        <Factory size={16} /> Confirmar Fabricación ({effectiveUnits} {selectedProd.unit}s)
                      </button>
                    </div>
                  </div>
                );
              })()}
            </form>
          </motion.div>
        </div>
      )}

      {/* Edit Product Modal */}
      {productToEdit && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-xl rounded-3xl p-6 md:p-8 border border-main shadow-2xl overflow-y-auto max-h-[92vh]"
          >
            <div className="flex items-center justify-between pb-4 border-b border-main/20 mb-6">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-accent/10 border border-accent/20 flex items-center justify-center text-accent">
                  <Pencil size={20} />
                </div>
                <div>
                  <h3 className="text-xl font-black uppercase tracking-tight text-main">Editar Producto</h3>
                  <p className="text-muted text-[11px] font-bold">Modifica nombre, tipo de inventario, unidad y stock</p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setProductToEdit(null)}
                className="p-2 text-muted hover:text-main hover:bg-main/5 rounded-xl transition-all"
                title="Cerrar"
              >
                <X size={20} />
              </button>
            </div>

            <div className="mb-6 p-3 rounded-xl bg-main/5 border border-main/10 flex items-center justify-between text-xs">
              <span className="text-muted font-bold">Stock en almacén:</span>
              <span className="font-black text-accent text-sm">
                {productStock[productToEdit.id] ?? 0} {productToEdit.unit}s
              </span>
            </div>

            <form onSubmit={handleSaveEditProduct} className="space-y-5">
              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Nombre del Producto *
                </label>
                <input 
                  type="text" 
                  required
                  placeholder="ej. Harina de Trigo, Leche Pasteurizada..."
                  className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main focus:outline-none focus:border-accent text-sm"
                  value={editProductForm.name}
                  onChange={e => setEditProductForm({ ...editProductForm, name: e.target.value })}
                />
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Tipo de Inventario *
                </label>
                <div className="grid grid-cols-3 gap-2">
                  {(['materia prima', 'producción', 'producto terminado'] as const).map((t) => (
                    <button
                      key={t}
                      type="button"
                      onClick={() => setEditProductForm({ ...editProductForm, type: t })}
                      className={cn(
                        "text-[10px] font-black uppercase py-2.5 px-2 rounded-xl border transition-all text-center flex flex-col items-center justify-center gap-0.5",
                        editProductForm.type === t
                          ? "bg-accent text-white border-accent shadow-md shadow-brand-gold-shadow"
                          : "bg-sidebar text-muted border-main hover:border-accent/40 hover:text-main"
                      )}
                    >
                      <span className="tracking-wider">{t}</span>
                      {editProductForm.type === t && <span className="text-[8px] opacity-80 font-bold">✓ Seleccionado</span>}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Unidad de Medida *
                </label>
                <div className="grid grid-cols-3 gap-2">
                  {[
                    { id: 'litro' as const, label: 'Litro (L)' },
                    { id: 'unidad' as const, label: 'Unidad (Und)' },
                    { id: 'kilo' as const, label: 'Kilo (Kg)' }
                  ].map((u) => (
                    <button
                      key={u.id}
                      type="button"
                      onClick={() => setEditProductForm({ ...editProductForm, unit: u.id })}
                      className={cn(
                        "text-[10px] font-black uppercase py-2.5 px-2 rounded-xl border transition-all text-center flex flex-col items-center justify-center gap-0.5",
                        editProductForm.unit === u.id
                          ? "bg-accent text-white border-accent shadow-md shadow-brand-gold-shadow"
                          : "bg-sidebar text-muted border-main hover:border-accent/40 hover:text-main"
                      )}
                    >
                      <span className="tracking-wider">{u.label}</span>
                      {editProductForm.unit === u.id && <span className="text-[8px] opacity-80 font-bold">✓ Activo</span>}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                    Categoría *
                  </label>
                  <input 
                    type="text" 
                    required
                    placeholder="ej. Insumos, Lácteos, Envases..."
                    className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main focus:outline-none focus:border-accent text-xs"
                    value={editProductForm.category}
                    onChange={e => setEditProductForm({ ...editProductForm, category: e.target.value })}
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                    Stock Mínimo (Alerta) *
                  </label>
                  <input 
                    type="number" 
                    required 
                    min="0"
                    className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main focus:outline-none focus:border-accent text-xs"
                    value={editProductForm.minStock}
                    onFocus={e => e.target.select()}
                    onChange={e => setEditProductForm({ ...editProductForm, minStock: Number(e.target.value) })}
                  />
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Descripción (Opcional)
                </label>
                <input 
                  type="text"
                  placeholder="Detalles adicionales del producto o especificaciones..."
                  className="w-full bg-sidebar border border-main rounded-xl px-4 py-3 font-bold text-main focus:outline-none focus:border-accent text-xs"
                  value={editProductForm.description}
                  onChange={e => setEditProductForm({ ...editProductForm, description: e.target.value })}
                />
              </div>

              {/* Consumo de Materiales / Receta en Edición */}
              {(editProductForm.type === 'producto terminado' || editProductForm.type === 'producción') && (
                <div className="p-4 rounded-2xl bg-amber-500/5 border border-amber-500/20 space-y-3">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2 text-amber-500">
                      <Layers size={16} />
                      <span className="text-[10px] font-black uppercase tracking-wider">
                        Consumo de Materiales (Fórmula / Receta)
                      </span>
                    </div>
                    <span className="text-[9px] font-bold text-muted uppercase">
                      Por 1 {editProductForm.unit}
                    </span>
                  </div>
                  <p className="text-[10px] text-muted font-medium">
                    Especifica qué insumos y materias primas consume este producto para fabricarse:
                  </p>

                  {/* Lista de materiales actuales */}
                  {editProductForm.recipe.length > 0 && (
                    <div className="space-y-1.5">
                      {editProductForm.recipe.map((item, idx) => {
                        const mat = products.find(p => p.id === item.materialProductId);
                        return (
                          <div key={idx} className="flex items-center justify-between p-2 rounded-xl bg-sidebar border border-main text-xs">
                            <div className="flex items-center gap-2">
                              <span className="w-5 h-5 rounded-full bg-accent/10 text-accent font-black text-[10px] flex items-center justify-center">
                                {idx + 1}
                              </span>
                              <div>
                                <span className="font-bold text-main">{mat?.name || item.materialName || 'Material'}</span>
                                <span className="text-[9px] text-muted block uppercase tracking-wider">
                                  Consume: <b className="text-accent">{item.quantityNeeded}</b> {mat?.unit || item.unit || 'unidad'}s
                                </span>
                              </div>
                            </div>
                            <button
                              type="button"
                              onClick={() => setEditProductForm(prev => ({
                                ...prev,
                                recipe: prev.recipe.filter((_, i) => i !== idx)
                              }))}
                              className="p-1 text-red-500 hover:bg-red-500/10 rounded-lg transition-colors"
                              title="Quitar material"
                            >
                              <X size={14} />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {/* Añadir material a la receta en edición */}
                  <div className="flex flex-col sm:flex-row gap-2 pt-1">
                    <select
                      id="edit-recipe-material-select"
                      className="flex-[2] bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main outline-none"
                      defaultValue=""
                    >
                      <option value="" disabled>Seleccione materia prima...</option>
                      {products
                        .filter(p => p.id !== productToEdit.id && (p.type === 'materia prima' || p.type === 'producción'))
                        .map(p => (
                          <option key={p.id} value={p.id}>
                            {p.name} ({p.unit}s) - Stock: {productStock[p.id] || 0}
                          </option>
                        ))
                      }
                    </select>
                    <input
                      id="edit-recipe-qty-input"
                      type="number"
                      step="0.01"
                      min="0.001"
                      placeholder="Cantidad"
                      className="flex-1 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main outline-none w-24"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        const sel = document.getElementById('edit-recipe-material-select') as HTMLSelectElement;
                        const qty = document.getElementById('edit-recipe-qty-input') as HTMLInputElement;
                        if (!sel || !sel.value || !qty || Number(qty.value) <= 0) {
                          alert('Seleccione un material e ingrese la cantidad consumida.');
                          return;
                        }
                        const mat = products.find(p => p.id === sel.value);
                        setEditProductForm(prev => ({
                          ...prev,
                          recipe: [
                            ...prev.recipe.filter(r => r.materialProductId !== sel.value),
                            {
                              materialProductId: sel.value,
                              materialName: mat?.name,
                              quantityNeeded: Number(qty.value),
                              unit: mat?.unit
                            }
                          ]
                        }));
                        qty.value = '';
                      }}
                      className="bg-accent text-white px-3 py-2 rounded-xl font-bold text-xs hover:brightness-110 flex items-center justify-center gap-1"
                    >
                      <Plus size={14} /> Añadir
                    </button>
                  </div>
                </div>
              )}

              {/* Empaque (Caja o Bulto) en Edición */}
              {(editProductForm.type === 'producto terminado' || editProductForm.type === 'producción') && (
                <div className="p-4 rounded-2xl bg-blue-500/5 border border-blue-500/20 space-y-3">
                  <div className="flex items-center gap-2 text-blue-500">
                    <Boxes size={16} />
                    <span className="text-[10px] font-black uppercase tracking-wider">
                      Empaque (Caja o Bulto y cuántas unidades lleva)
                    </span>
                  </div>
                  <div className="grid grid-cols-4 gap-1.5">
                    {[
                      { id: 'ninguno' as const, label: 'Sin Empaque' },
                      { id: 'caja' as const, label: 'En Cajas' },
                      { id: 'bulto' as const, label: 'En Bultos' },
                      { id: 'ambos' as const, label: 'Cajas y Bultos' }
                    ].map(opt => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setEditProductForm(prev => ({
                          ...prev,
                          packaging: { ...prev.packaging, packagingType: opt.id }
                        }))}
                        className={cn(
                          "text-[9px] font-black uppercase py-2 px-1 rounded-xl border transition-all text-center",
                          editProductForm.packaging.packagingType === opt.id
                            ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                            : "bg-sidebar text-muted border-main hover:bg-main/5"
                        )}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>

                  {(editProductForm.packaging.packagingType === 'caja' || editProductForm.packaging.packagingType === 'ambos') && (
                    <div className="flex items-center justify-between gap-3 pt-1">
                      <label className="text-[10px] font-black text-muted uppercase tracking-wider flex items-center gap-1.5">
                        📦 ¿Cuántas unidades lleva cada <b>Caja</b>?
                      </label>
                      <input
                        type="number"
                        min="1"
                        placeholder="ej. 12"
                        value={editProductForm.packaging.unitsPerBox || ''}
                        onChange={e => setEditProductForm(prev => ({
                          ...prev,
                          packaging: { ...prev.packaging, unitsPerBox: Number(e.target.value) || 0 }
                        }))}
                        className="w-28 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main focus:border-accent outline-none text-right"
                      />
                    </div>
                  )}

                  {(editProductForm.packaging.packagingType === 'bulto' || editProductForm.packaging.packagingType === 'ambos') && (
                    <div className="flex items-center justify-between gap-3 pt-1">
                      <label className="text-[10px] font-black text-muted uppercase tracking-wider flex items-center gap-1.5">
                        🧺 ¿Cuántas unidades lleva cada <b>Bulto</b>?
                      </label>
                      <input
                        type="number"
                        min="1"
                        placeholder="ej. 50"
                        value={editProductForm.packaging.unitsPerBundle || ''}
                        onChange={e => setEditProductForm(prev => ({
                          ...prev,
                          packaging: { ...prev.packaging, unitsPerBundle: Number(e.target.value) || 0 }
                        }))}
                        className="w-28 bg-sidebar border border-main rounded-xl px-3 py-2 text-xs font-bold text-main focus:border-accent outline-none text-right"
                      />
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-3 pt-4 border-t border-main/20">
                <button 
                  type="button" 
                  onClick={() => setProductToEdit(null)}
                  className="flex-1 bg-main/5 text-muted py-3.5 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest"
                >
                  Cancelar
                </button>
                <button 
                  type="submit"
                  disabled={!editProductForm.name.trim()}
                  className="flex-[2] bg-accent text-white py-3.5 rounded-xl font-black tracking-widest hover:brightness-110 shadow-lg shadow-brand-gold-shadow uppercase text-xs disabled:opacity-50 transition-all flex items-center justify-center gap-2"
                >
                  <Check size={16} /> Guardar Cambios
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {productToDelete && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-sm rounded-3xl p-8 border border-red-500/30 shadow-2xl shadow-red-500/10 text-center"
          >
            <div className="w-16 h-16 bg-red-500/10 text-red-500 rounded-full flex items-center justify-center mx-auto mb-6">
              <Trash2 size={32} />
            </div>
            <h3 className="text-xl font-black mb-2 uppercase tracking-tight">Confirmar Eliminación</h3>
            <p className="text-muted text-sm mb-8">
              ¿Seguro que desea eliminar <span className="text-main font-bold">"{productToDelete.name}"</span>? 
              Se borrarán permanentemente todos sus lotes y el historial de movimientos.
            </p>
            <div className="flex gap-4">
              <button 
                onClick={() => setProductToDelete(null)}
                className="flex-1 bg-main/5 text-muted py-3 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-sm uppercase tracking-widest"
              >
                Cancelar
              </button>
              <button 
                onClick={confirmDeleteProduct}
                className="flex-[1.5] bg-red-500 text-white py-3 rounded-xl font-black hover:bg-red-600 transition-all shadow-lg shadow-red-500/20 text-sm uppercase tracking-widest"
              >
                Eliminar Ahora
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Modal Editar Compra */}
      {purchaseToEdit && (
        <div className="fixed inset-0 bg-main/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-lg rounded-3xl p-6 md:p-8 shadow-2xl border border-main overflow-y-auto max-h-[90vh]"
          >
            <div className="flex justify-between items-start mb-6">
              <div>
                <span className="text-[10px] font-black uppercase tracking-[0.2em] text-accent block mb-1">
                  Módulo de Compras
                </span>
                <h3 className="text-xl md:text-2xl font-black uppercase tracking-tight text-main">
                  Editar Registro de Compra
                </h3>
                <p className="text-xs text-muted font-bold mt-1 uppercase tracking-wider">
                  Producto: {products.find(p => p.id === purchaseToEdit.productId)?.name || 'Producto'} • Lote: {purchaseToEdit.id.slice(0, 8)}
                </p>
              </div>
              <button 
                onClick={() => setPurchaseToEdit(null)}
                className="p-2 text-muted hover:text-main rounded-xl hover:bg-main/5 transition-all"
              >
                <X size={20} />
              </button>
            </div>

            <form onSubmit={handleSaveEditPurchase} className="space-y-4">
              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Fecha de la Compra
                </label>
                <input 
                  type="date" required
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                  value={editPurchaseForm.entryDate}
                  onChange={e => setEditPurchaseForm({ ...editPurchaseForm, entryDate: e.target.value })}
                />
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Proveedor
                </label>
                <input 
                  type="text" required
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                  value={editPurchaseForm.supplier}
                  onChange={e => setEditPurchaseForm({ ...editPurchaseForm, supplier: e.target.value })}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <div className="flex justify-between items-center mb-2">
                    <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em]">
                      Cantidad
                    </label>
                    {purchaseToEdit.remainingQuantity < purchaseToEdit.quantity && (
                      <span className="text-[9px] font-black text-amber-500 uppercase">
                        Consumido: {(purchaseToEdit.quantity - purchaseToEdit.remainingQuantity).toFixed(1)}
                      </span>
                    )}
                  </div>
                  <input 
                    type="number" step="0.01" min="0.01" required
                    className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main font-mono text-xs"
                    value={editPurchaseForm.quantity}
                    onChange={e => setEditPurchaseForm({ ...editPurchaseForm, quantity: Number(e.target.value) })}
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                    Costo Unitario (Bs.)
                  </label>
                  <input 
                    type="number" step="0.01" min="0" required
                    className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main font-mono text-xs"
                    value={editPurchaseForm.unitCost}
                    onChange={e => setEditPurchaseForm({ ...editPurchaseForm, unitCost: Number(e.target.value) })}
                  />
                </div>
              </div>

              <div className="p-3 bg-accent/5 rounded-xl border border-accent/20 flex justify-between items-center">
                <span className="text-[10px] font-black uppercase tracking-wider text-muted">
                  Total de la Compra:
                </span>
                <span className="text-sm font-black text-accent font-mono">
                  Bs. {(Number(editPurchaseForm.quantity) * Number(editPurchaseForm.unitCost)).toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </span>
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Descripción / Factura / Notas
                </label>
                <textarea 
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs min-h-[70px]"
                  value={editPurchaseForm.description}
                  onChange={e => setEditPurchaseForm({ ...editPurchaseForm, description: e.target.value })}
                  placeholder="Número de factura o notas de compra..."
                />
              </div>

              <div className="flex gap-3 pt-3 border-t border-main/20">
                <button 
                  type="button"
                  onClick={() => {
                    const buy = purchaseToEdit;
                    setPurchaseToEdit(null);
                    setPurchaseToDelete(buy);
                  }}
                  className="p-3 bg-red-500/10 text-red-500 rounded-xl hover:bg-red-500 hover:text-white transition-all border border-red-500/20"
                  title="Eliminar esta compra"
                >
                  <Trash2 size={16} />
                </button>
                <button 
                  type="button"
                  onClick={() => setPurchaseToEdit(null)}
                  className="flex-1 bg-main/5 text-muted py-3 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest"
                >
                  Cancelar
                </button>
                <button 
                  type="submit"
                  className="flex-[2] bg-accent text-white py-3 rounded-xl font-black tracking-widest hover:brightness-110 shadow-lg shadow-brand-gold-shadow uppercase text-xs transition-all flex items-center justify-center gap-2"
                >
                  <Check size={16} /> Guardar Cambios
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Modal Editar Venta */}
      {saleToEdit && (
        <div className="fixed inset-0 bg-main/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-lg rounded-3xl p-6 md:p-8 shadow-2xl border border-main overflow-y-auto max-h-[90vh]"
          >
            <div className="flex justify-between items-start mb-6">
              <div>
                <span className="text-[10px] font-black uppercase tracking-[0.2em] text-accent block mb-1">
                  Módulo de Ventas
                </span>
                <h3 className="text-xl md:text-2xl font-black uppercase tracking-tight text-main">
                  Editar Registro de Venta
                </h3>
                <p className="text-xs text-muted font-bold mt-1 uppercase tracking-wider">
                  Producto: {products.find(p => p.id === saleToEdit.productId)?.name || 'Producto'} • Movimiento: {saleToEdit.id.slice(0, 8)}
                </p>
              </div>
              <button 
                onClick={() => setSaleToEdit(null)}
                className="p-2 text-muted hover:text-main rounded-xl hover:bg-main/5 transition-all"
              >
                <X size={20} />
              </button>
            </div>

            <form onSubmit={handleSaveEditSale} className="space-y-4">
              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Fecha de la Venta
                </label>
                <input 
                  type="date" required
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main uppercase text-xs"
                  value={editSaleForm.date}
                  onChange={e => setEditSaleForm({ ...editSaleForm, date: e.target.value })}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                    Cantidad Vendida
                  </label>
                  <input 
                    type="number" step="0.01" min="0.01" required
                    className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main font-mono text-xs"
                    value={editSaleForm.quantity}
                    onChange={e => setEditSaleForm({ ...editSaleForm, quantity: Number(e.target.value) })}
                  />
                </div>

                <div>
                  <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                    Precio Venta Unitario (Bs.)
                  </label>
                  <input 
                    type="number" step="0.01" min="0" required
                    className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main font-mono text-xs"
                    value={editSaleForm.salePrice}
                    onChange={e => setEditSaleForm({ ...editSaleForm, salePrice: Number(e.target.value) })}
                  />
                </div>
              </div>

              <div className="p-3 bg-accent/5 rounded-xl border border-accent/20 flex justify-between items-center">
                <span className="text-[10px] font-black uppercase tracking-wider text-muted">
                  Total de la Venta Facturada:
                </span>
                <span className="text-sm font-black text-accent font-mono">
                  Bs. {(Number(editSaleForm.quantity) * Number(editSaleForm.salePrice)).toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </span>
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Factura / Referencia / Comprobante
                </label>
                <input 
                  type="text"
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs"
                  value={editSaleForm.reference}
                  onChange={e => setEditSaleForm({ ...editSaleForm, reference: e.target.value })}
                  placeholder="Número de factura o referencia..."
                />
              </div>

              <div>
                <label className="block text-[10px] font-black text-muted uppercase tracking-[0.3em] mb-2">
                  Notas / Cliente / Descripción
                </label>
                <textarea 
                  className="w-full bg-black/5 dark:bg-black/30 border border-main rounded-xl px-5 py-3 focus:outline-none focus:border-accent font-bold text-main text-xs min-h-[70px]"
                  value={editSaleForm.description}
                  onChange={e => setEditSaleForm({ ...editSaleForm, description: e.target.value })}
                  placeholder="Cliente o notas de la venta..."
                />
              </div>

              <div className="flex gap-3 pt-3 border-t border-main/20">
                <button 
                  type="button"
                  onClick={() => {
                    const s = saleToEdit;
                    setSaleToEdit(null);
                    setSaleToDelete(s);
                  }}
                  className="p-3 bg-red-500/10 text-red-500 rounded-xl hover:bg-red-500 hover:text-white transition-all border border-red-500/20"
                  title="Eliminar esta venta"
                >
                  <Trash2 size={16} />
                </button>
                <button 
                  type="button"
                  onClick={() => setSaleToEdit(null)}
                  className="flex-1 bg-main/5 text-muted py-3 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest"
                >
                  Cancelar
                </button>
                <button 
                  type="submit"
                  className="flex-[2] bg-accent text-white py-3 rounded-xl font-black tracking-widest hover:brightness-110 shadow-lg shadow-brand-gold-shadow uppercase text-xs transition-all flex items-center justify-center gap-2"
                >
                  <Check size={16} /> Guardar Cambios
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Modal Confirmar Eliminación de Compra */}
      {purchaseToDelete && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-md rounded-3xl p-6 md:p-8 border border-red-500/30 shadow-2xl shadow-red-500/10 text-center"
          >
            <div className="w-16 h-16 bg-red-500/10 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-5 border border-red-500/20">
              <Trash2 size={30} />
            </div>
            <span className="text-[10px] font-black uppercase tracking-[0.2em] text-red-500 block mb-1">
              Módulo de Compras
            </span>
            <h3 className="text-xl font-black mb-3 uppercase tracking-tight text-main">
              ¿Eliminar Registro de Compra?
            </h3>
            
            {(() => {
              const prod = products.find(p => p.id === purchaseToDelete.productId);
              const prodName = prod?.name || 'Producto';
              const unit = prod?.unit || 'unidad';
              const consumed = Math.max(0, purchaseToDelete.quantity - purchaseToDelete.remainingQuantity);
              const totalAmount = purchaseToDelete.quantity * purchaseToDelete.unitCost;

              return (
                <div className="text-left bg-main/5 p-4 rounded-2xl border border-main/10 mb-6 space-y-2 text-xs">
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Producto:</span>
                    <span className="font-black text-main uppercase">{prodName}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Proveedor:</span>
                    <span className="font-bold text-main">{purchaseToDelete.supplier}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Fecha Compra:</span>
                    <span className="font-medium text-main">{format(new Date(purchaseToDelete.entryDate), 'dd/MM/yyyy')}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Cantidad Comprada:</span>
                    <span className="font-black text-accent">{purchaseToDelete.quantity} {unit}s</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Costo Total:</span>
                    <span className="font-mono font-bold text-main">Bs. {totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                  </div>

                  {consumed > 0 && (
                    <div className="mt-3 p-3 rounded-xl bg-amber-500/10 border border-amber-500/30 text-amber-600 dark:text-amber-400">
                      <div className="flex items-center gap-1.5 font-black text-[11px] mb-1">
                        <AlertTriangle size={14} /> ADVERTENCIA DE EXISTENCIA
                      </div>
                      <p className="text-[10px] leading-relaxed">
                        Ya se han consumido o vendido <b>{consumed.toFixed(2)}</b> {unit}s de este lote. Quedan <b>{purchaseToDelete.remainingQuantity.toFixed(2)}</b> {unit}s en almacén que se descontarán del inventario.
                      </p>
                    </div>
                  )}
                </div>
              );
            })()}

            <div className="flex gap-3">
              <button 
                type="button"
                onClick={() => setPurchaseToDelete(null)}
                disabled={isDeletingMovement}
                className="flex-1 bg-main/5 text-muted py-3.5 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest disabled:opacity-50"
              >
                Cancelar
              </button>
              <button 
                type="button"
                onClick={executeConfirmDeletePurchase}
                disabled={isDeletingMovement}
                className="flex-[1.5] bg-red-500 text-white py-3.5 rounded-xl font-black hover:bg-red-600 transition-all shadow-lg shadow-red-500/20 text-xs uppercase tracking-widest disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isDeletingMovement ? (
                  <span>Eliminando...</span>
                ) : (
                  <>
                    <Trash2 size={15} />
                    <span>Eliminar Compra</span>
                  </>
                )}
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Modal Confirmar Eliminación en Lote de Compras */}
      {bulkPurchasesToDelete && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-md rounded-3xl p-6 md:p-8 border border-red-500/30 shadow-2xl shadow-red-500/10 text-center"
          >
            <div className="w-16 h-16 bg-red-500/10 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-5 border border-red-500/20">
              <Trash2 size={30} />
            </div>
            <span className="text-[10px] font-black uppercase tracking-[0.2em] text-red-500 block mb-1">
              Eliminación Masiva
            </span>
            <h3 className="text-xl font-black mb-3 uppercase tracking-tight text-main">
              ¿Eliminar {bulkPurchasesToDelete.length} Compras?
            </h3>
            <p className="text-muted text-xs mb-6">
              Esta acción eliminará de forma permanente los <b>{bulkPurchasesToDelete.length} registros</b> de compra seleccionados y actualizará las existencias del inventario.
            </p>

            <div className="flex gap-3">
              <button 
                type="button"
                onClick={() => setBulkPurchasesToDelete(null)}
                disabled={isDeletingMovement}
                className="flex-1 bg-main/5 text-muted py-3.5 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest disabled:opacity-50"
              >
                Cancelar
              </button>
              <button 
                type="button"
                onClick={executeConfirmDeleteBulkPurchases}
                disabled={isDeletingMovement}
                className="flex-[1.5] bg-red-500 text-white py-3.5 rounded-xl font-black hover:bg-red-600 transition-all shadow-lg shadow-red-500/20 text-xs uppercase tracking-widest disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isDeletingMovement ? 'Eliminando...' : `Eliminar ${bulkPurchasesToDelete.length} Compras`}
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Modal Confirmar Eliminación de Venta */}
      {saleToDelete && (
        <div className="fixed inset-0 bg-main/90 backdrop-blur-xl z-[60] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card w-full max-w-md rounded-3xl p-6 md:p-8 border border-red-500/30 shadow-2xl shadow-red-500/10 text-center"
          >
            <div className="w-16 h-16 bg-red-500/10 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-5 border border-red-500/20">
              <Trash2 size={30} />
            </div>
            <span className="text-[10px] font-black uppercase tracking-[0.2em] text-red-500 block mb-1">
              Módulo de Ventas
            </span>
            <h3 className="text-xl font-black mb-3 uppercase tracking-tight text-main">
              ¿Eliminar Registro de Venta?
            </h3>

            {(() => {
              const prod = products.find(p => p.id === saleToDelete.productId);
              const prodName = prod?.name || 'Producto';
              const unit = prod?.unit || 'unidad';
              const unitPrice = saleToDelete.salePrice || saleToDelete.unitCost || 0;
              const totalAmount = saleToDelete.quantity * unitPrice;

              return (
                <div className="text-left bg-main/5 p-4 rounded-2xl border border-main/10 mb-6 space-y-2 text-xs">
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Producto:</span>
                    <span className="font-black text-main uppercase">{prodName}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Cantidad Vendida:</span>
                    <span className="font-black text-accent">{saleToDelete.quantity} {unit}s</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Precio Unitario:</span>
                    <span className="font-mono font-bold text-main">Bs. {unitPrice.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Total Facturado:</span>
                    <span className="font-mono font-black text-accent">Bs. {totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted font-bold">Fecha de Venta:</span>
                    <span className="font-medium text-main">{format(new Date(saleToDelete.date), 'dd/MM/yyyy HH:mm')}</span>
                  </div>
                  {(saleToDelete.reference || saleToDelete.description) && (
                    <div className="flex justify-between">
                      <span className="text-muted font-bold">Referencia/Nota:</span>
                      <span className="font-medium text-main truncate max-w-[200px]">{saleToDelete.reference || saleToDelete.description}</span>
                    </div>
                  )}

                  {/* Restitución de Stock Info Box */}
                  <div className="mt-3 p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/30 text-emerald-600 dark:text-emerald-400">
                    <div className="flex items-center gap-1.5 font-black text-[11px] mb-1">
                      <PackageCheck size={14} /> RESTITUCIÓN AUTOMÁTICA DE STOCK
                    </div>
                    <p className="text-[10px] leading-relaxed">
                      Al anular esta venta, las <b>{saleToDelete.quantity} {unit}s</b> serán <b>devueltas automáticamente al inventario físico</b> y el Kardex valorado se actualizará en tiempo real.
                    </p>
                  </div>
                </div>
              );
            })()}

            <div className="flex gap-3">
              <button 
                type="button"
                onClick={() => setSaleToDelete(null)}
                disabled={isDeletingMovement}
                className="flex-1 bg-main/5 text-muted py-3.5 rounded-xl font-bold border border-main hover:bg-main/10 transition-all text-xs uppercase tracking-widest disabled:opacity-50"
              >
                Cancelar
              </button>
              <button 
                type="button"
                onClick={executeConfirmDeleteSale}
                disabled={isDeletingMovement}
                className="flex-[1.5] bg-red-500 text-white py-3.5 rounded-xl font-black hover:bg-red-600 transition-all shadow-lg shadow-red-500/20 text-xs uppercase tracking-widest disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isDeletingMovement ? (
                  <span>Restituyendo...</span>
                ) : (
                  <>
                    <Trash2 size={15} />
                    <span>Eliminar y Restituir</span>
                  </>
                )}
              </button>
            </div>
          </motion.div>
        </div>
      )}
      {/* Toast System */}
      <div className="fixed bottom-8 right-8 z-[200] flex flex-col gap-4 pointer-events-none">
        <AnimatePresence>
          {toasts.map(toast => (
            <motion.div 
              key={toast.id}
              initial={{ opacity: 0, x: 100, scale: 0.9 }}
              animate={{ opacity: 1, x: 0, scale: 1 }}
              exit={{ opacity: 0, x: 100, scale: 0.9 }}
              className="pointer-events-auto bg-card border border-main rounded-2xl shadow-2xl p-4 flex items-center gap-4 w-[320px] relative overflow-hidden"
            >
              <div className={cn(
                "w-1 h-full absolute left-0 top-0",
                toast.type === 'LOW_STOCK' ? "bg-red-500" : "bg-green-500"
              )} />
              <div className={cn(
                "w-10 h-10 rounded-xl flex items-center justify-center shrink-0",
                toast.type === 'LOW_STOCK' ? "bg-red-500/10 text-red-500" : "bg-green-500/10 text-green-500"
              )}>
                {toast.type === 'LOW_STOCK' ? <AlertTriangle size={20} /> : <CheckCircle size={20} />}
              </div>
              <div className="flex-1">
                <p className="text-[10px] font-black uppercase tracking-widest text-main mb-0.5">{toast.title}</p>
                <p className="text-[9px] font-bold text-muted leading-tight">{toast.message}</p>
              </div>
              <button 
                onClick={() => setToasts(prev => prev.filter(t => t.id !== toast.id))}
                className="text-muted hover:text-main p-1 transition-colors"
              >
                <X size={14} />
              </button>
            </motion.div>
          ))}
        </AnimatePresence>
      </div>

      <NotificationCenter />
    </div>
  );
}

