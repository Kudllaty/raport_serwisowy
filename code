// System Raportów Serwisowych - logika aplikacji

// Funkcja wywoływana przy publikacji jako aplikacja webowa
function doGet() {
  return HtmlService.createHtmlOutputFromFile('OrderPanel')
    .setTitle('System Raportów Serwisowych')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)  // Zmiana z NATIVE na IFRAME
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')  // Dodanie metatagu viewport
    .setFaviconUrl('https://www.gstatic.com/script/apps_script_1x.png');
}

// Funkcja wywoływana przy otwarciu arkusza
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Raporty Serwisowe')
    .addItem('Otwórz panel raportów', 'showOrderPanel')
    .addToUi();
}

// Otwiera panel raportów
function showOrderPanel() {
  var html = HtmlService.createHtmlOutputFromFile('OrderPanel')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('System Raportów Serwisowych');
  SpreadsheetApp.getUi().showModalDialog(html, 'System Raportów Serwisowych');
}

// Funkcja logowania
function login(login, password) {
  try {
    const users = getDataFromSheet('Użytkownicy');
    const user = users.find(u => u.Login === login && u.Hasło === password);
    
    if (user) {
      return {
        success: true,
        serwis: user.Serwis,
        rola: user.Rola
      };
    } else {
      return { success: false, message: "Nieprawidłowy login lub hasło" };
    }
  } catch(e) {
    return { success: false, message: "Błąd logowania: " + e.message };
  }
}

// Pobiera wszystkie serwisy
function getServices() {
  try {
    const services = getDataFromSheet('Serwisy');
    const uniqueServices = [...new Set(services.map(s => s.Serwis))];
    return { success: true, services: uniqueServices };
  } catch(e) {
    return { success: false, message: "Błąd pobierania serwisów: " + e.message };
  }
}

// Pobiera kategorie napraw dla wybranego serwisu
function getRepairCategories(serwis) {
  try {
    const services = getDataFromSheet('Serwisy');
    const filtered = services.filter(s => s.Serwis === serwis);
    const uniqueCategories = [...new Set(filtered.map(s => s['Kategoria naprawy']))];
    return { success: true, categories: uniqueCategories };
  } catch(e) {
    return { success: false, message: "Błąd pobierania kategorii napraw: " + e.message };
  }
}

// Pobiera operatorów dla wybranego serwisu i kategorii naprawy
function getOperators(serwis, category) {
  try {
    const services = getDataFromSheet('Serwisy');
    const filtered = services.filter(s => s.Serwis === serwis && s['Kategoria naprawy'] === category);
    const uniqueOperators = [...new Set(filtered.map(s => s.Operator))];
    return { success: true, operators: uniqueOperators };
  } catch(e) {
    return { success: false, message: "Błąd pobierania operatorów: " + e.message };
  }
}

// Pobiera oddziały dla wybranego serwisu, kategorii naprawy i operatora
function getDepartments(serwis, category, operator) {
  try {
    const services = getDataFromSheet('Serwisy');
    const filtered = services.filter(s => 
      s.Serwis === serwis && 
      s['Kategoria naprawy'] === category && 
      s.Operator === operator
    );
    const uniqueDepartments = [...new Set(filtered.map(s => s.Oddział))];
    return { success: true, departments: uniqueDepartments };
  } catch(e) {
    return { success: false, message: "Błąd pobierania oddziałów: " + e.message };
  }
}

// Pobiera adresy dla wybranego serwisu
function getAddresses(serwis) {
  try {
    const addresses = getDataFromSheet('Adresy');
    const filtered = addresses.filter(a => a.Serwis === serwis);
    return { success: true, addresses: filtered };
  } catch(e) {
    return { success: false, message: "Błąd pobierania adresów: " + e.message };
  }
}

// Pobiera dane adresowe dla serwisu
function getServiceAddressData(serwis) {
  const addresses = getDataFromSheet('Adresy');
  const address = addresses.find(a => a.Serwis === serwis);
  if (address) {
    return {
      ulica: address.Ulica || '',
      kodPocztowy: address['Kod pocztowy'] || '',
      miasto: address.Miasto || '',
      email: address.Email || '',
      telefon: address.Telefon || ''
    };
  }
  return {
    ulica: '',
    kodPocztowy: '',
    miasto: '',
    email: '',
    telefon: ''
  };
}

// Funkcja klienta do pobierania danych adresowych - dostępna z JavaScript
function getAddressDataForClient(serwis) {
  // Po prostu wywołuje istniejącą funkcję pomocniczą
  return getServiceAddressData(serwis);
}

// Zapisuje nowy adres
function saveAddress(address) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Adresy');
    
    // Dodaj nowy wiersz
    sheet.appendRow([
      address.serwis,
      address.kodPocztowy,
      address.miasto,
      address.ulica,
      address.email,
      address.telefon
    ]);
    
    return { 
      success: true, 
      message: "Adres został dodany pomyślnie",
      address: address
    };
  } catch(e) {
    return { success: false, message: "Błąd zapisywania adresu: " + e.message };
  }
}

// Edytuje istniejący adres
function updateAddress(oldAddress, newAddress) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Adresy');
    const data = sheet.getDataRange().getValues();
    
    // Normalize property names to handle different case formats
    const oldSerwis = oldAddress.serwis || oldAddress.Serwis || '';
    const oldKodPocztowy = oldAddress.kodPocztowy || oldAddress['Kod pocztowy'] || '';
    const oldMiasto = oldAddress.miasto || oldAddress.Miasto || '';
    const oldUlica = oldAddress.ulica || oldAddress.Ulica || '';
    
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][0] === oldSerwis &&
        data[i][1] === oldKodPocztowy &&
        data[i][2] === oldMiasto &&
        data[i][3] === oldUlica
      ) {
        sheet.getRange(i + 1, 2).setValue(newAddress.kodPocztowy);
        sheet.getRange(i + 1, 3).setValue(newAddress.miasto);
        sheet.getRange(i + 1, 4).setValue(newAddress.ulica);
        sheet.getRange(i + 1, 5).setValue(newAddress.email);
        sheet.getRange(i + 1, 6).setValue(newAddress.telefon);
        break;
      }
    }
    
    return { 
      success: true, 
      message: "Adres został zaktualizowany pomyślnie",
      address: newAddress
    };
  } catch(e) {
    return { success: false, message: "Błąd aktualizacji adresu: " + e.message };
  }
}

// Usuwa adres
function deleteAddress(address) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Adresy');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][0] === address.serwis &&
        data[i][1] === address.kodPocztowy &&
        data[i][2] === address.miasto &&
        data[i][3] === address.ulica
      ) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    return { 
      success: true, 
      message: "Adres został usunięty pomyślnie"
    };
  } catch(e) {
    return { success: false, message: "Błąd usuwania adresu: " + e.message };
  }
}

// Pobiera Cennik ID na podstawie danych podstawowych
function getPriceListId(serwis, category, operator, department) {
  try {
    const services = getDataFromSheet('Serwisy');
    const match = services.find(s => 
      s.Serwis === serwis && 
      s['Kategoria naprawy'] === category && 
      s.Operator === operator &&
      s.Oddział === department
    );
    
    if (match) {
      return { success: true, cennikId: match['Cennik ID'] };
    } else {
      return { success: false, message: "Nie znaleziono pasującego cennika" };
    }
  } catch(e) {
    return { success: false, message: "Błąd pobierania ID cennika: " + e.message };
  }
}

// Wyszukuje produkty na podstawie Cennik ID i kategorii naprawy
function searchProducts(cennikId, category, searchTerm = '') {
  try {
    const products = getDataFromSheet('Cennik części');
    
    // Dodaj logowanie do łatwiejszej diagnostyki
    console.log(`Wyszukiwanie produktów dla: Cennik ID=${cennikId}, Kategoria naprawy=${category}`);
    console.log(`Liczba wszystkich produktów w arkuszu: ${products.length}`);
    
    // Sprawdź, jakie nazwy pól faktycznie istnieją w produktach
    if (products.length > 0) {
      console.log(`Nazwy pól w pierwszym produkcie: ${Object.keys(products[0]).join(', ')}`);
    }
    
    // Używaj bardziej elastycznego filtrowania, które uwzględni różne możliwe nazwy pól
    let filtered = products.filter(p => {
      // Sprawdź różne możliwe nazwy pola "Cennik ID"
      const cennikIdMatch = 
        (p['Cennik ID'] === cennikId) || 
        (p['cennik ID'] === cennikId) || 
        (p['cennik id'] === cennikId) || 
        (p['CENNIK ID'] === cennikId);
      
      // Sprawdź różne możliwe nazwy pola "Kategoria naprawy" 
      const categoryMatch = 
        (p['Kategoria naprawy'] === category) || 
        (p['kategoria naprawy'] === category) || 
        (p['Kategoria Naprawy'] === category) || 
        (p['KATEGORIA NAPRAWY'] === category);
      
      return cennikIdMatch && categoryMatch;
    });
    
    console.log(`Liczba przefiltrowanych produktów przed szukaniem tekstu: ${filtered.length}`);
    
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(p => {
        // Sprawdź czy pola Kod i Nazwa części istnieją w danym produkcie
        const kodMatch = p.Kod && p.Kod.toString().toLowerCase().includes(term);
        const nazwaMatch = (p['Nazwa części'] && p['Nazwa części'].toString().toLowerCase().includes(term)) || 
                          (p['nazwa części'] && p['nazwa części'].toString().toLowerCase().includes(term));
        
        return kodMatch || nazwaMatch;
      });
      console.log(`Liczba produktów po zastosowaniu wyszukiwania tekstu "${searchTerm}": ${filtered.length}`);
    }
    
    // Wypisz kilka pierwszych znalezionych produktów do celów diagnostycznych
    if (filtered.length > 0) {
      console.log(`Przykładowe znalezione produkty: ${JSON.stringify(filtered.slice(0, 2))}`);
    } else {
      console.log(`Nie znaleziono żadnych produktów dla kombinacji: Cennik ID=${cennikId}, Kategoria naprawy=${category}`);
    }
    
    return { success: true, products: filtered };
  } catch(e) {
    console.error(`Błąd wyszukiwania produktów: ${e.message}`);
    console.error(`Stos błędu: ${e.stack}`);
    return { success: false, message: "Błąd wyszukiwania produktów: " + e.message };
  }
}

// Pobiera kategorie produktów na podstawie Cennik ID i kategorii naprawy
function getProductCategories(cennikId, category) {
  try {
    const products = getDataFromSheet('Cennik części');
    
    // Logowanie dla celów diagnostycznych
    console.log(`Pobieranie kategorii produktów dla: Cennik ID=${cennikId}, Kategoria naprawy=${category}`);
    console.log(`Liczba wszystkich produktów w arkuszu: ${products.length}`);
    
    // Używaj bardziej elastycznego filtrowania, które uwzględni różne możliwe nazwy pól
    let filtered = products.filter(p => {
      // Sprawdź różne możliwe nazwy pola "Cennik ID"
      const cennikIdMatch = 
        (p['Cennik ID'] === cennikId) || 
        (p['cennik ID'] === cennikId) || 
        (p['cennik id'] === cennikId) || 
        (p['CENNIK ID'] === cennikId);
      
      // Sprawdź różne możliwe nazwy pola "Kategoria naprawy" 
      const categoryMatch = 
        (p['Kategoria naprawy'] === category) || 
        (p['kategoria naprawy'] === category) || 
        (p['Kategoria Naprawy'] === category) || 
        (p['KATEGORIA NAPRAWY'] === category);
      
      return cennikIdMatch && categoryMatch;
    });
    
    console.log(`Liczba przefiltrowanych produktów: ${filtered.length}`);
    
    // Pobierz unikalne kategorie produktów
    const productCategories = [];
    
    // Sprawdź różne możliwe nazwy pola "Kategoria produktu"
    filtered.forEach(p => {
      let kategoria = null;
      if (p['Kategoria produktu']) kategoria = p['Kategoria produktu'];
      else if (p['kategoria produktu']) kategoria = p['kategoria produktu'];
      else if (p['KATEGORIA PRODUKTU']) kategoria = p['KATEGORIA PRODUKTU'];
      
      if (kategoria && !productCategories.includes(kategoria)) {
        productCategories.push(kategoria);
      }
    });
    
    // Posortuj kategorie alfabetycznie
    productCategories.sort();
    
    console.log(`Znalezione kategorie produktów (${productCategories.length}): ${productCategories.join(', ')}`);
    
    return { success: true, categories: productCategories };
  } catch(e) {
    console.error(`Błąd pobierania kategorii produktów: ${e.message}`);
    console.error(`Stos błędu: ${e.stack}`);
    return { success: false, message: "Błąd pobierania kategorii produktów: " + e.message };
  }
}

/**
 * Funkcja do wyszukiwania części w zakładce "Cennik części"
 * @param {string} cennikId - ID cennika
 * @param {string} category - Kategoria naprawy
 * @param {string} searchTerm - Fraza do wyszukania (opcjonalna)
 * @return {Object} Obiekt zawierający znalezione części lub informację o błędzie
 */
function searchPartsInModal(cennikId, category, searchTerm = '') {
  try {
    const parts = getDataFromSheet('Cennik części');
    
    // Logowanie dla celów diagnostycznych
    Logger.log(`Wyszukiwanie części dla: Cennik ID=${cennikId}, Kategoria naprawy=${category}`);
    Logger.log(`Liczba wszystkich części w arkuszu: ${parts.length}`);
    
    // Sprawdź strukturę danych części
    if (parts.length > 0) {
      Logger.log(`Struktura pierwszej części: ${JSON.stringify(parts[0])}`);
    }
    
    // Elastyczne filtrowanie uwzględniające różne możliwe nazwy pól
    let filtered = parts.filter(p => {
      // Obsługa różnych możliwych nazw pola "Cennik ID"
      const cennikIdMatch = 
        (p['Cennik ID'] === cennikId) || 
        (p['cennik ID'] === cennikId) || 
        (p['cennik id'] === cennikId) || 
        (p['CENNIK ID'] === cennikId);
      
      // Obsługa różnych możliwych nazw pola "Kategoria naprawy"
      const categoryMatch = 
        (p['Kategoria naprawy'] === category) || 
        (p['kategoria naprawy'] === category) || 
        (p['Kategoria Naprawy'] === category) || 
        (p['KATEGORIA NAPRAWY'] === category);
      
      return cennikIdMatch && categoryMatch;
    });
    
    Logger.log(`Liczba przefiltrowanych części przed wyszukiwaniem: ${filtered.length}`);
    
    // Filtrowanie po frazie wyszukiwania jeśli podana
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(p => {
        // Sprawdź czy pola Kod i Nazwa części istnieją w danej części
        const kodMatch = p.Kod && p.Kod.toString().toLowerCase().includes(term);
        const nazwaMatch = (p['Nazwa części'] && p['Nazwa części'].toString().toLowerCase().includes(term)) || 
                           (p['nazwa części'] && p['nazwa części'].toString().toLowerCase().includes(term));
        
        return kodMatch || nazwaMatch;
      });
      Logger.log(`Liczba części po zastosowaniu wyszukiwania "${searchTerm}": ${filtered.length}`);
    }
    
    // Diagnostyka znalezionych części
    if (filtered.length > 0) {
      Logger.log(`Przykładowe znalezione części: ${JSON.stringify(filtered.slice(0, 2))}`);
    } else {
      Logger.log(`Nie znaleziono części dla: Cennik ID=${cennikId}, Kategoria=${category}`);
    }
    
    return { success: true, parts: filtered };
  } catch(e) {
    Logger.log(`Błąd wyszukiwania części: ${e.message}`);
    Logger.log(`Stos błędu: ${e.stack}`);
    return { success: false, message: "Błąd wyszukiwania części: " + e.message };
  }
}

// Funkcje do obsługi wyszukiwania usług
function searchServices(cennikId, category, searchTerm = '') {
  try {
    const services = getDataFromSheet('Cennik usługi');
    let filtered = services.filter(s => 
      s['Cennik ID'] === cennikId && 
      s['Kategoria naprawy'] === category
    );
    
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(s => 
        s.Kod.toLowerCase().includes(term) || 
        s['Nazwa usługi'].toLowerCase().includes(term)
      );
    }
    
    return { success: true, services: filtered };
  } catch(e) {
    return { success: false, message: "Błąd wyszukiwania usług: " + e.message };
  }
}

// Funkcja wyszukująca ramy rowerów
function searchFrameNumbers(department) {
  try {
    // Pobieramy dane z zakładki "Rowery"
    const frames = getDataFromSheet('Rowery');
    console.log(`Znaleziono ${frames.length} rowerów w zakładce Rowery`);
    console.log(`Filtrowanie rowerów dla oddziału: ${department}`);
    
    // Sprawdzenie dostępnych pól w arkuszu "Rowery"
    if (frames.length > 0) {
      console.log("Dostępne pola w arkuszu Rowery:", Object.keys(frames[0]).join(", "));
    }
    
    // Filtrujemy tylko te rowery, które należą do wybranego oddziału
    let filtered = frames;
    if (department) {
      filtered = frames.filter(f => f.Oddział === department);
    }
    
    console.log(`Po filtrowaniu znaleziono ${filtered.length} rowerów dla oddziału ${department}`);
    
    // Formatowanie dat ostatniego przeglądu i przekształcanie do formatu JSON z potrzebnymi danymi
    const frameNumbers = filtered.map(frame => {
      // Przygotowanie daty ostatniego przeglądu w formacie ISO jeśli istnieje
      let lastInspectionDate = null;
      if (frame['Ostatni przegląd']) {
        if (frame['Ostatni przegląd'] instanceof Date) {
          lastInspectionDate = Utilities.formatDate(
            frame['Ostatni przegląd'], 
            Session.getScriptTimeZone(), 
            'yyyy-MM-dd'
          );
        } else if (typeof frame['Ostatni przegląd'] === 'string') {
          // Próba konwersji ze stringa jeśli to nie jest obiekt Date
          try {
            const dateParts = frame['Ostatni przegląd'].split('.');
            if (dateParts.length === 3) {
              const date = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`);
              lastInspectionDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } else {
              lastInspectionDate = frame['Ostatni przegląd']; // Użyj jak jest jeśli nie udało się przekształcić
            }
          } catch (e) {
            console.error("Błąd konwersji daty:", e);
            lastInspectionDate = frame['Ostatni przegląd']; // Użyj jak jest jeśli nie udało się przekształcić
          }
        }
      }
      
      // Tworzenie obiektu JSON z danymi roweru, który będzie pokazywany w liście wyboru
      // oraz używany do wypełnienia formularza
      const frameData = {
        serialNumber: frame['Nr ramy'] || '',
        lastInspection: lastInspectionDate,
        model: frame['Model'] || '',
        addedDate: frame['Data dodania'] instanceof Date ? 
          Utilities.formatDate(frame['Data dodania'], Session.getScriptTimeZone(), 'yyyy-MM-dd') : 
          (frame['Data dodania'] || '')
      };
      
      // Łańcuch opisowy dla select option pokazujący numer ramy i datę ostatniego przeglądu
      const displayText = `${frameData.serialNumber} ${lastInspectionDate ? '(ostatni przegląd: ' + lastInspectionDate + ')' : ''}`;
      
      return {
        value: JSON.stringify(frameData), // Serializujemy dane roweru do JSON
        text: displayText           // Tekst wyświetlany w selectboxie
      };
    });
    
    console.log(`Przygotowano ${frameNumbers.length} pozycji dla listy wyboru numerów ram`);
    
    return { success: true, frames: frameNumbers };
  } catch(e) {
    console.error("Błąd wyszukiwania numerów ram:", e);
    return { success: false, message: "Błąd wyszukiwania numerów ram: " + e.message };
  }
}

// Funkcja zapisująca nowy numer ramy
function saveNewFrameNumber(frameData) {
  try {
    // Sprawdzenie wymaganych pól
    if (!frameData.serialNumber || !frameData.department) {
      return { 
        success: false, 
        message: "Brakujące dane: numer ramy i oddział są wymagane" 
      };
    }
    
    // Pobieranie danych z arkusza Rowery
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Rowery');
    
    if (!sheet) {
      return {
        success: false,
        message: "Nie znaleziono arkusza 'Rowery'"
      };
    }
    
    // Pobierz wszystkie dane z arkusza
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Pierwszy wiersz to nagłówki
    
    // Znajdź indeksy kolumn dla oddziału i numeru ramy
    const departmentColIndex = headers.indexOf('Oddział');
    const serialNumberColIndex = headers.indexOf('Nr ramy');
    
    if (departmentColIndex === -1 || serialNumberColIndex === -1) {
      return {
        success: false,
        message: "Nie znaleziono wymaganych kolumn w arkuszu 'Rowery'"
      };
    }
    
    // Sprawdź czy istnieje już taka kombinacja Oddział+Nr ramy
    for (let i = 1; i < data.length; i++) {
      if (data[i][departmentColIndex] === frameData.department && 
          data[i][serialNumberColIndex] === frameData.serialNumber) {
        return {
          success: false,
          message: `Numer ramy ${frameData.serialNumber} dla oddziału ${frameData.department} już istnieje w bazie`
        };
      }
    }
    
    // Jeśli dotarliśmy tutaj, to znaczy że nie ma duplikatu, możemy dodać nowy rekord
    const currentDate = new Date();
    
    // Przygotowanie wiersza do dodania
    // Struktura: Oddział, Nr ramy, Ostatni przegląd, Model, Data dodania
    const newRow = [];
    
    // Wypełnij wszystkie kolumny (nawet jeśli niektóre będą puste)
    for (let i = 0; i < headers.length; i++) {
      switch(headers[i]) {
        case 'Oddział':
          newRow[i] = frameData.department;
          break;
        case 'Nr ramy':
          newRow[i] = frameData.serialNumber;
          break;
        case 'Model':
          newRow[i] = frameData.model || '';
          break;
        case 'Data dodania':
          newRow[i] = currentDate;
          break;
        case 'Ostatni przegląd':
          newRow[i] = ''; // Pusty, bo to nowy numer ramy
          break;
        default:
          newRow[i] = '';  // Domyślne puste wartości dla pozostałych kolumn
          break;
      }
    }
    
    // Dodaj nowy wiersz do arkusza
    sheet.appendRow(newRow);
    
    // Przygotuj dane do zwrócenia
    const returnFrameData = {
      serialNumber: frameData.serialNumber,
      model: frameData.model || '',
      department: frameData.department,
      addedDate: Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    };
    
    // Zwróć sukces i dane
    return {
      success: true,
      message: `Nowy numer ramy ${frameData.serialNumber} został dodany do oddziału ${frameData.department}`,
      frameData: returnFrameData
    };
    
  } catch (e) {
    console.error("Błąd podczas zapisywania numeru ramy:", e);
    return {
      success: false,
      message: "Błąd podczas zapisywania numeru ramy: " + e.message
    };
  }
}

// Zapisuje zamówienie
function saveOrder(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Generuj unikalne ID zamówienia
    const orderId = generateOrderId();
    const currentDate = new Date();
    
    // Poprawione pobieranie danych adresowych - obsługa różnych formatów pól
    const adresUlica = orderData.address.ulica || orderData.address.Ulica || '';
    const adresKodPocztowy = orderData.address.kodPocztowy || orderData.address['Kod pocztowy'] || '';
    const adresMiasto = orderData.address.miasto || orderData.address.Miasto || '';
    const adresEmail = orderData.address.email || orderData.address.Email || '';
    const adresTelefon = orderData.address.telefon || orderData.address.Telefon || '';
    
    // Zapewnienie, że orderData.items jest zdefiniowane i jest tablicą
    if (!orderData.items || !Array.isArray(orderData.items)) {
      orderData.items = [];
    }
    
    // Oblicz łączną ilość artykułów
    const totalItemsCount = orderData.items.reduce((total, item) => total + item.quantity, 0);
    
    // Zapisz główne dane zamówienia
    const orderSheet = ss.getSheetByName('Zamówienia');
    orderSheet.appendRow([
      orderId,
      currentDate,
      orderData.serwis,
      orderData.category,
      orderData.operator,
      orderData.department,
      `${adresUlica}, ${adresKodPocztowy} ${adresMiasto}`, // Poprawiony adres
      adresEmail,
      adresTelefon,
      orderData.cennikId,
      calculateTotalValue(orderData.items),
      totalItemsCount, // Dodana nowa kolumna "Ilość artykułów"
      'Nowe',         // Status
      '',             // Uwagi
      orderData.notes,
      currentDate     // Data aktualizacji
    ]);
    
    // Zapisz elementy zamówienia z nowymi nazwami kolumn
    const itemsSheet = ss.getSheetByName('Elementy zamówień');
    orderData.items.forEach(item => {
      itemsSheet.appendRow([
        orderId,         // ID zamówienia
        orderData.serwis,// Serwis
        item.category,   // Kategoria produktu
        orderData.operator, // Operator
        orderData.department, // Oddział
        item.code,       // Kod produktu
        item.name,       // Nazwa
        item.price,      // Cena
        item.quantity,   // Ilość
        item.price * item.quantity, // Wartość
        currentDate      // Data dodania
      ]);
    });
    
    // Wyślij maile potwierdzające
    sendConfirmationEmails(orderData, orderId);
    
    return { 
      success: true, 
      message: "Zamówienie zostało zapisane pomyślnie. Numer zamówienia: " + orderId,
      orderId: orderId 
    };
  } catch(e) {
    return { success: false, message: "Błąd zapisywania zamówienia: " + e.message };
  }
}

// Generuje unikalne ID zamówienia
function generateOrderId() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const dateStr = `${year}${month}${day}`;
  
  // Pobranie ostatniego numeru zamówienia i inkrementacja - niezależnie od daty
  const lastOrderNum = getLastOrderNumber(dateStr) || 0;
  const newOrderNum = lastOrderNum + 1;
  const paddedOrderNum = String(newOrderNum).padStart(5, '0'); // Format 00001
  
  console.log(`Generowanie nowego ID zamówienia: ZAM/${dateStr}/${paddedOrderNum} (poprzedni numer: ${lastOrderNum})`);
  return `ZAM/${dateStr}/${paddedOrderNum}`;
}

// Funkcja do pobierania ostatniego numeru zamówienia z całego arkusza
function getLastOrderNumber(dateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Zamówienia');
    const data = sheet.getDataRange().getValues();
    let lastNum = 0;
    
    // Przeszukaj wszystkie zamówienia, nie tylko te z bieżącego dnia
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0]; // Zakładając, że ID jest w pierwszej kolumnie
      if (id && id.includes('ZAM/')) {
        const parts = id.split('/');
        if (parts.length === 3) {
          const num = parseInt(parts[2]);
          if (!isNaN(num) && num > lastNum) {
            lastNum = num;
          }
        }
      }
    }
    
    console.log(`Znaleziono ostatni numer zamówienia: ${lastNum}`);
    return lastNum;
  } catch(e) {
    console.error("Błąd pobierania ostatniego numeru zamówienia:", e);
    return 0; // Awaryjnie zwróć 0, co spowoduje start od 00001
  }
}

// Oblicza łączną wartość zamówienia
function calculateTotalValue(items) {
  return items.reduce((total, item) => total + (item.price * item.quantity), 0);
}

// Wysyła maile potwierdzające zamówienie
function sendConfirmationEmails(orderData, orderId) {
  // Mail do serwisu (bez cen)
  sendMailToService(orderData, orderId);
  
  // Mail do opiekuna (z cenami)
  sendMailToSupervisor(orderData, orderId);
}

// Zapisuje raport serwisowy
function saveReport(reportData) {
  try {
    // Walidacja danych reportData
    if (!reportData) {
      return { success: false, message: "Brak danych raportu" };
    }
    
    // Zapewnienie, że reportData.parts i reportData.services są zdefiniowane i są tablicami
    if (!reportData.parts || !Array.isArray(reportData.parts)) {
      reportData.parts = [];
    }
    
    if (!reportData.services || !Array.isArray(reportData.services)) {
      reportData.services = [];
    }
    
    // Pobierz dane adresowe z zakładki Adresy na podstawie wybranego serwisu
    const addressData = getServiceAddressData(reportData.serwis);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Raporty serwisowe');
    
    // Generuj unikalne ID raportu
    const reportId = generateReportId();
    const currentDate = new Date();
    
    // Używamy danych adresowych z pobranego obiektu addressData
    const adresUlica = addressData.ulica;
    const adresKodPocztowy = addressData.kodPocztowy;
    const adresMiasto = addressData.miasto;
    const adresEmail = addressData.email;
    const adresTelefon = addressData.telefon;

    // Bezpieczne odczytywanie danych
    const serviceDate = reportData.serviceDate ? new Date(reportData.serviceDate) : currentDate;
    const serviceType = reportData.serviceType || '';
    const bikeMileage = reportData.bikeMileage || '';
    let laborHours = reportData.laborHours;
    if (typeof laborHours === 'string' && laborHours.includes('.')) {
      laborHours = laborHours.replace('.', ',');
    }
    const isOnSite = reportData.isOnSite || 'Nie';
    
    // 1. Zapisz do głównej tabeli "Raport serwisowy"
    const reportSheet = ss.getSheetByName('Raport serwisowy');
    if (!reportSheet) {
      // Jeśli arkusz nie istnieje, utwórz go
      const newSheet = ss.insertSheet('Raport serwisowy');
      // Dodaj nagłówki BEZ Adres, Email, Telefon
      newSheet.appendRow([
        'ID raportu', 'Data utworzenia', 'Serwis', 'Kategoria naprawy', 'Operator', 'Oddział',
        'Numer ramy', 'Rodzaj serwisu', 'Data serwisu', 'Przebieg (km)',
        'Roboczo godziny', 'Dojazd', 'Wartość usług', 'Wartość części', 'Wartość całkowita', 'Status',
        'Link do zdjęć', 'Uwagi', 'Data aktualizacji'
      ]);
    }
    
    // Pobierz aktualne ceny z cennika części dla wybranej kombinacji Cennik ID i kategorii naprawy
    const cennikParts = getPartsFromPriceList(reportData.cennikId, reportData.category);
    console.log(`Znaleziono ${cennikParts.length} części w cenniku dla ID=${reportData.cennikId} i kategorii=${reportData.category}`);
    
    // Aktualizuj ceny części na podstawie cennika
    const partItems = reportData.parts.map(part => {
      // Znajdź tę część w cenniku po kodzie
      const cennikPart = cennikParts.find(p => p.Kod === part.code);
      
      if (cennikPart && cennikPart.Cena) {
        // Jeśli znaleziono część i ma cenę, użyj jej
        part.price = cennikPart.Cena;
        console.log(`Zaktualizowano cenę dla części ${part.code} (${part.name}) na ${part.price}zł z cennika`);
      } else {
        // Jeśli nie znaleziono, zachowaj obecną lub ustaw na 0 jeśli nie istnieje
        if (!part.price) {
          console.warn(`Nie znaleziono ceny dla części ${part.code} (${part.name}) w cenniku!`);
          part.price = 0;
        }
      }
      
      return part;
    });
    
    // Obliczanie wartości usług, części i całkowitej
    const servicesValue = reportData.services.reduce((sum, service) => sum + (parseFloat(service.price) || 0), 0);
    const partsValue = partItems.reduce((sum, part) => sum + ((parseFloat(part.price) || 0) * (parseInt(part.quantity) || 1)), 0);
    const totalValue = servicesValue + partsValue;
    
    // Zapisz wiersz BEZ Adres, Email, Telefon
    reportSheet.appendRow([
      reportId,                // ID raportu
      currentDate,             // Data utworzenia
      reportData.serwis,       // Serwis
      reportData.category,     // Kategoria naprawy
      reportData.operator,     // Operator
      reportData.department,   // Oddział
      reportData.bikeSerialNumber || '', // Numer ramy
      serviceType,             // Rodzaj serwisu
      serviceDate,             // Data serwisu
      bikeMileage,             // Przebieg (km)
      laborHours,              // Roboczo godziny
      isOnSite,                // Dojazd
      servicesValue,           // Wartość usług
      partsValue,              // Wartość części
      totalValue,              // Wartość całkowita
      'Nowy',                  // Status
      '',                      // Link do zdjęć (uzupełniany później)
      reportData.notes || '',  // Uwagi
      currentDate              // Data aktualizacji
    ]);
    
    // 2a. Podział elementów raportu - usługi
    const serviceItems = reportData.services || [];
    
    // Zapisz usługi do "Elementy raportów usługi"
    const serviceItemsSheet = ss.getSheetByName('Elementy raportów usługi');
    if (!serviceItemsSheet) {
      // Jeśli arkusz nie istnieje, utwórz go
      const newSheet = ss.insertSheet('Elementy raportów usługi');
      // Dodaj nagłówki dokładnie w takiej kolejności jak w wymaganiach
      newSheet.appendRow([
        'Elementy raportów usługi', 'Serwis', 'Kategoria', 'Operator', 'Oddział', 'Numer ramy', 
        'Kod usługi', 'Nazwa usługi', 'Cena', 'Niestandardowa', 'Data dodania'
      ]);
    }
    
    serviceItems.forEach(item => {
      // Dodano numer ramy do tabeli Elementy raportów usługi
      serviceItemsSheet.appendRow([
        reportId,         // ID raportu
        reportData.serwis,// Serwis
        item.category || reportData.category, // Kategoria
        reportData.operator, // Operator
        reportData.department, // Oddział
        reportData.bikeSerialNumber || '', // Numer ramy
        item.code,       // Kod usługi
        item.name,       // Nazwa usługi
        item.price,      // Cena
        item.isCustom ? 'Tak' : 'Nie', // Niestandardowa
        currentDate      // Data dodania
      ]);
    });
    
    // 2b. Podział elementów raportu - części
    
    // Zapisz części do "Elementy raportów części" z zaktualizowanymi cenami
    const partItemsSheet = ss.getSheetByName('Elementy raportów części');
    if (!partItemsSheet) {
      // Jeśli arkusz nie istnieje, utwórz go
      const newSheet = ss.insertSheet('Elementy raportów części');
      // Dodaj nagłówki dokładnie w takiej kolejności jak w wymaganiach
      newSheet.appendRow([
        'Elementy raportów części', 'Serwis', 'Kategoria', 'Operator', 'Oddział', 'Numer ramy', 
        'Kod części', 'Nazwa części', 'Cena', 'Ilość', 'Wartość', 'Data dodania'
      ]);
    }
    
    partItems.forEach(item => {
      const itemPrice = item.price || 0;
      const quantity = item.quantity || 1;
      const value = itemPrice * quantity;
      
      partItemsSheet.appendRow([
        reportId,         // ID raportu
        reportData.serwis,// Serwis
        item.category || reportData.category, // Kategoria
        reportData.operator, // Operator
        reportData.department, // Oddział
        reportData.bikeSerialNumber || '', // Numer ramy - dodany brakujący element
        item.code,       // Kod części
        item.name,       // Nazwa części
        itemPrice,      // Cena
        quantity,   // Ilość
        value, // Wartość
        currentDate      // Data dodania
      ]);
    });
    
    // Aktualizacja reportData z poprawionymi cenami przed wysłaniem maili
    reportData.parts = partItems;
    
    // Wyślij maile potwierdzające
    sendReportConfirmationEmails(reportData, reportId);
    
    return { 
      success: true, 
      message: "Raport serwisowy został zapisany pomyślnie. Numer raportu: " + reportId,
      reportId: reportId 
    };
  } catch(e) {
    console.error("Błąd zapisywania raportu serwisowego:", e);
    return { success: false, message: "Błąd zapisywania raportu: " + e.message };
  }
}

// Pobiera części z cennika dla danego ID cennika i kategorii naprawy
function getPartsFromPriceList(cennikId, category) {
  try {
    const parts = getDataFromSheet('Cennik części');
    
    // Filtruj części według ID cennika i kategorii naprawy
    const filtered = parts.filter(p => {
      const cennikIdMatch = 
        (p['Cennik ID'] === cennikId) || 
        (p['cennik ID'] === cennikId) || 
        (p['cennik id'] === cennikId);
      
      const categoryMatch = 
        (p['Kategoria naprawy'] === category) || 
        (p['kategoria naprawy'] === category) || 
        (p['Kategoria Naprawy'] === category);
      
      return cennikIdMatch && categoryMatch;
    });
    
    // Loguj znalezione części do debugowania
    if (filtered.length > 0) {
      console.log(`Znaleziono ${filtered.length} części w cenniku dla ID=${cennikId} i kategorii=${category}`);
    } else {
      console.warn(`Nie znaleziono części w cenniku dla ID=${cennikId} i kategorii=${category}`);
    }
    
    return filtered;
  } catch(e) {
    console.error(`Błąd pobierania części z cennika: ${e.message}`);
    return []; // Zwróć pustą tablicę w przypadku błędu
  }
}

// Funkcja do pobierania ostatniego numeru raportu
function getLastReportNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Raport serwisowy');
    
    if (!sheet) {
      return 0; // Jeśli arkusz nie istnieje, zwróć 0
    }
    
    const data = sheet.getDataRange().getValues();
    let lastNum = 0;
    
    // Przeszukaj wszystkie raporty
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0]; // Zakładając, że ID jest w pierwszej kolumnie
      if (id && id.includes('SER/')) {
        const parts = id.split('/');
        if (parts.length === 3) {
          const num = parseInt(parts[2]);
          if (!isNaN(num) && num > lastNum) {
            lastNum = num;
          }
        }
      }
    }
    
    console.log(`Znaleziono ostatni numer raportu: ${lastNum}`);
    return lastNum;
  } catch(e) {
    console.error("Błąd pobierania ostatniego numeru raportu:", e);
    return 0; // Awaryjnie zwróć 0, co spowoduje start od 00001
  }
}

// Generuje unikalne ID raportu serwisowego
function generateReportId() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const dateStr = `${year}${month}${day}`;
  
  // Pobranie ostatniego numeru raportu i inkrementacja - niezależnie od daty
  const lastReportNum = getLastReportNumber() || 0;
  const newReportNum = lastReportNum + 1;
  const paddedReportNum = String(newReportNum).padStart(5, '0'); // Format 00001
  
  console.log(`Generowanie nowego ID raportu: SER/${dateStr}/${paddedReportNum} (poprzedni numer: ${lastReportNum})`);
  return `SER/${dateStr}/${paddedReportNum}`;
}

// Tworzy arkusz raportów serwisowych jeśli nie istnieje
function createReportsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Raporty serwisowe');
    
    if (!sheet) {
      sheet = ss.insertSheet('Raporty serwisowe');
      
      // Ustaw nagłówki
      sheet.getRange(1, 1, 1, 20).setValues([
        [
          'ID raportu', 
          'Data utworzenia', 
          'Serwis', 
          'Kategoria naprawy', 
          'Operator', 
          'Oddział', 
          'Dane klienta', 
          'Email klienta', 
          'Telefon klienta', 
          'Model roweru', 
          'Numer seryjny', 
          'Data serwisu', 
          'Status naprawy', 
          'Opis problemu', 
          'Wykonane czynności', 
          'Link zdjęć przed', 
          'Link zdjęć po', 
          'Ilość części', 
          'Uwagi', 
          'Data aktualizacji'
        ]
      ]);
      
      // Formatowanie
      sheet.getRange(1, 1, 1, 20).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, 20);
    }
    
    // Sprawdź też czy istnieją arkusze dla elementów raportów
    let elementsSheet = ss.getSheetByName('Elementy raportów');
    if (!elementsSheet) {
      elementsSheet = ss.insertSheet('Elementy raportów');
      
      // Ustaw nagłówki
      elementsSheet.getRange(1, 1, 1, 11).setValues([
        [
          'ID raportu', 
          'Serwis', 
          'Kategoria produktu', 
          'Operator', 
          'Oddział', 
          'Kod produktu', 
          'Nazwa produktu', 
          'Cena', 
          'Ilość', 
          'Wartość', 
          'Data dodania'
        ]
      ]);
      
      // Formatowanie
      elementsSheet.getRange(1, 1, 1, 11).setFontWeight('bold');
      elementsSheet.setFrozenRows(1);
      elementsSheet.autoResizeColumns(1, 11);
    }
    
    return true;
  } catch(e) {
    console.error("Błąd podczas tworzenia arkusza raportów:", e);
    return false;
  }
}

// Wysyła maile potwierdzające raport serwisowy
function sendReportConfirmationEmails(reportData, reportId) {
  console.log("Wysyłanie maili potwierdzających dla raportu ID:", reportId);
  
  // Dodanie ID raportu do obiektu reportData, aby był dostępny w funkcjach mailowych
  reportData.id = reportId;
  
  // Poprawione mapowanie pól services/parts na serviceItems/partItems
  reportData.serviceItems = reportData.services || [];
  reportData.partItems = reportData.parts || [];
  
  console.log("Liczba usług w raporcie:", reportData.serviceItems.length);
  console.log("Liczba części w raporcie:", reportData.partItems.length);
  
  // Przygotowanie danych dla maili
  const emailData = {
    id: reportId,
    serwis: reportData.serwis,
    category: reportData.category,
    operator: reportData.operator,
    department: reportData.department,
    serviceType: reportData.serviceType || 'Nie określono',
    serviceDate: reportData.serviceDate ? new Date(reportData.serviceDate).toLocaleDateString('pl-PL') : new Date().toLocaleDateString('pl-PL'),
    bikeMileage: reportData.bikeMileage || 'Nie określono',
    laborHours: reportData.laborHours || 'Nie określono',
    isOnSite: reportData.isOnSite || 'Nie',
    bikeModel: reportData.bikeModel || '',
    bikeSerialNumber: reportData.bikeSerialNumber || '',
    notes: reportData.notes || '',
    email: reportData.address?.email || '',
    phone: reportData.address?.telefon || '',
    serviceItems: reportData.serviceItems,
    partItems: reportData.partItems
  };
  
  console.log("Przygotowane dane do maila:", JSON.stringify(emailData));
  
  // Mail do serwisu (bez cen)
  const clientEmailStatus = sendReportEmailToClient(reportData);
  console.log("Status wysyłania maila do serwisu:", clientEmailStatus);
  
  // Mail do opiekuna serwisu (z cenami)
  const supervisorEmailStatus = sendMailToSupervisor(reportData);
  console.log("Status wysyłania maila do opiekuna:", supervisorEmailStatus);
  
  return { success: (clientEmailStatus || supervisorEmailStatus) };
}

// Pobiera dane z arkusza
function getDataFromSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Arkusz "${sheetName}" nie istnieje`);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    console.log(`Nagłówki w arkuszu ${sheetName}:`, headers);
    const rows = data.slice(1);
    
    let result = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    console.log(`Dane z arkusza ${sheetName}:`, JSON.stringify(result));
    return result;
  } catch(e) {
    console.error(`Błąd pobierania danych z arkusza ${sheetName}:`, e);
    
    // Jeśli wystąpi błąd, zwróć dane testowe
    return getMockData(sheetName);
  }
}

// Konwertuje URL z Google Drive na miniaturę
function extractFileId(driveUrl) {
  // Format: https://drive.google.com/file/d/FILE_ID/view
  const match = driveUrl.match(/\/d\/([^\/]+)/);
  if (match && match[1]) {
    return match[1];
  }
  
  // Format: https://drive.google.com/open?id=FILE_ID
  const idParam = driveUrl.match(/[?&]id=([^&]+)/);
  if (idParam && idParam[1]) {
    return idParam[1];
  }
  
  return null;
}

// Pobiera obrazy z arkusza "Logo" - poprawiony do używania "Zdjęcia 2"
function getAppImages() {
  try {
    // Spróbuj pobrać dane z arkusza "Logo"
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Logo');
    
    if (!sheet) {
      throw new Error("Nie znaleziono arkusza 'Logo'");
    }
    
    // Dokładne pobieranie obrazów z komórek w odpowiedniej kolejności
    // Zmodyfikowano kod, aby pobierać obrazy z kolumny C (Zdjęcia 2) zamiast kolumny B (Zdjęcia)
    const images = {
      logoA2: getCellImageUrl(sheet, 'A2'),  // Logo główne (bez zmian)
      logoB2: getCellImageUrl(sheet, 'C2'),  // Obraz 1 (pobierany z Zdjęcia 2)
      logoB3: getCellImageUrl(sheet, 'C3'),  // Obraz 2 (pobierany z Zdjęcia 2)
      logoB4: getCellImageUrl(sheet, 'C4'),  // Obraz 3 (pobierany z Zdjęcia 2)
      logoB5: getCellImageUrl(sheet, 'C5')   // Obraz 4 (pobierany z Zdjęcia 2)
    };
    
    // Dodaj dokładne informacje do dziennika
    console.log("Pobrane obrazy z arkusza Logo (z kolumny Zdjęcia 2):");
    console.log("- Logo główne (A2):", images.logoA2);
    console.log("- Obraz 1 (C2):", images.logoB2);
    console.log("- Obraz 2 (C3):", images.logoB3);
    console.log("- Obraz 3 (C4):", images.logoB4);
    console.log("- Obraz 4 (C5):", images.logoB5);
    
    return images;
  } catch(e) {
    console.error("Błąd pobierania obrazów:", e);
    // Awaryjne obrazy zastępcze
    return {
      logoA2: 'https://via.placeholder.com/200x80?text=CARGO+BIKE+SERWIS',
      logoB2: 'https://via.placeholder.com/300x200?text=Obraz+1',
      logoB3: 'https://via.placeholder.com/300x200?text=Obraz+2',
      logoB4: 'https://via.placeholder.com/300x200?text=Obraz+3', 
      logoB5: 'https://via.placeholder.com/300x200?text=Obraz+4'
    };
  }
}

// Pobiera URL obrazu z określonej komórki - poprawiona wersja z lepszym logowaniem
function getCellImageUrl(sheet, cellAddress) {
  try {
    console.log(`Pobieranie obrazu z komórki ${cellAddress}...`);
    const cell = sheet.getRange(cellAddress);
    const formulas = cell.getFormulas();
    const value = cell.getValue();
    
    console.log(`Komórka ${cellAddress} - formuła:`, formulas.length > 0 ? formulas[0][0] : "brak");
    console.log(`Komórka ${cellAddress} - wartość:`, value);
    
    // Sprawdź czy komórka zawiera formułę IMAGE()
    if (formulas.length > 0 && formulas[0].length > 0) {
      const formula = formulas[0][0];
      const urlMatch = formula.match(/=IMAGE\("([^"]+)"/);
      if (urlMatch && urlMatch[1]) {
        console.log(`Znaleziono URL w formule IMAGE() w komórce ${cellAddress}:`, urlMatch[1]);
        return urlMatch[1];
      }
    }
    
    // Sprawdź czy wartość komórki to URL
    if (typeof value === 'string' && (value.startsWith('http') || value.startsWith('/'))) {
      console.log(`Znaleziono URL jako wartość w komórce ${cellAddress}:`, value);
      return value;
    }
    
    // Sprawdź czy komórka zawiera adnotację z URL
    const note = cell.getNote();
    if (note && (note.startsWith('http') || note.startsWith('/'))) {
      console.log(`Znaleziono URL w adnotacji komórki ${cellAddress}:`, note);
      return note;
    }
    
    console.log(`Nie znaleziono URL w komórce ${cellAddress}`);
    return null;
  } catch(e) {
    console.error(`Błąd pobierania URL obrazu z komórki ${cellAddress}:`, e);
    return null;
  }
}

// Zwraca dane testowe do celów rozwojowych
function getMockData(sheetName) {
  const mockData = {
    'Użytkownicy': [
      { Login: 'test', Hasło: 'test', Serwis: 'Serwis A', Rola: 'User' },
      { Login: 'admin', Hasło: 'admin', Serwis: 'Serwis B', Rola: 'Admin' }
    ],
    'Serwisy': [
      { Serwis: 'Serwis A', Oddział: 'Oddział 1', Operator: 'Operator A', 'Kategoria naprawy': 'Kategoria 1', 'Cennik ID': 'C001' },
      { Serwis: 'Serwis A', Oddział: 'Oddział 2', Operator: 'Operator A', 'Kategoria naprawy': 'Kategoria 1', 'Cennik ID': 'C002' },
      { Serwis: 'Serwis A', Oddział: 'Oddział 1', Operator: 'Operator B', 'Kategoria naprawy': 'Kategoria 1', 'Cennik ID': 'C003' },
      { Serwis: 'Serwis A', Oddział: 'Oddział 1', Operator: 'Operator A', 'Kategoria naprawy': 'Kategoria 2', 'Cennik ID': 'C004' }
    ],
    'Adresy': [
      { Serwis: 'Serwis A', 'Kod pocztowy': '00-001', Miasto: 'Warszawa', Ulica: 'ul. Przykładowa 1' },
      { Serwis: 'Serwis B', 'Kod pocztowy': '00-002', Miasto: 'Kraków', Ulica: 'ul. Testowa 2' }
    ],
    'Cennik części': [
      { 'Cennik ID': 'C001', Kod: 'ADAPTER-Z39', 'Nazwa części': 'Adapter tarczy Z39', Cena: 199.99, 'Kategoria naprawy': 'Kategoria 1', 'Link do obrazu': 'https://drive.google.com/file/d/1sample/view' },
      { 'Cennik ID': 'C001', Kod: 'PART-X100', 'Nazwa części': 'Część X100', Cena: 299.99, 'Kategoria naprawy': 'Kategoria 1', 'Link do obrazu': 'https://example.com/image.jpg' },
      { 'Cennik ID': 'C002', Kod: 'TOOL-Y200', 'Nazwa części': 'Narzędzie Y200', Cena: 499.99, 'Kategoria naprawy': 'Kategoria 1', 'Link do obrazu': 'https://drive.google.com/file/d/2sample/view' }
    ],
    'Opiekun': [
      { Imię: 'Jan Kowalski', Email: 'opiekun@example.com' }
    ]
  };
  
  return mockData[sheetName] || [];
}

// Pobiera dane adresowe dla serwisu
function getServiceAddressData(serwis) {
  const addresses = getDataFromSheet('Adresy');
  const address = addresses.find(a => a.Serwis === serwis);
  if (address) {
    return {
      ulica: address.Ulica || '',
      kodPocztowy: address['Kod pocztowy'] || '',
      miasto: address.Miasto || '',
      email: address.Email || '',
      telefon: address.Telefon || ''
    };
  }
  return {
    ulica: '',
    kodPocztowy: '',
    miasto: '',
    email: '',
    telefon: ''
  };
}

// --- Funkcje obsługi zdjęć serwisowych ---

// Pobiera lub tworzy główny folder na zdjęcia serwisowe
function getOrCreateServiceImagesFolder() {
  try {
    // Adres głównego folderu ze zdjęciami serwisowymi
    const MAIN_FOLDER_NAME = "Zdjęcia Serwisowe CBS";
    
    // Sprawdź czy folder główny już istnieje
    const folderIterator = DriveApp.getFoldersByName(MAIN_FOLDER_NAME);
    
    if (folderIterator.hasNext()) {
      // Jeśli folder istnieje, zwróć go
      const folder = folderIterator.next();
      console.log("Znaleziono istniejący folder na zdjęcia serwisowe:", folder.getId());
      return folder;
    } else {
      // Jeśli nie istnieje, utwórz nowy folder w katalogu głównym
      const mainFolder = DriveApp.createFolder(MAIN_FOLDER_NAME);
      console.log("Utworzono nowy główny folder na zdjęcia serwisowe:", mainFolder.getId());
      return mainFolder;
    }
  } catch(e) {
    console.error("Błąd podczas pobierania folderu zdjęć serwisowych:", e);
    throw new Error("Nie można uzyskać dostępu do folderu zdjęć serwisowych: " + e.message);
  }
}

// Tworzy strukturę folderów dla danego raportu serwisowego
function createReportImagesFolders(reportId) {
  try {
    const mainFolder = getOrCreateServiceImagesFolder();
    
    // Tworzenie podfolderów według hierarchii: RaportID
    // Format reportId to np. SER/20250502/00001
    const reportFolderName = reportId.replace(/\//g, '_');
    
    // Sprawdź czy folder już istnieje
    const existingFolders = mainFolder.getFoldersByName(reportFolderName);
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    
    // Jeśli nie istnieje, utwórz nowy folder
    const reportFolder = mainFolder.createFolder(reportFolderName);
    console.log("Utworzono folder dla raportu:", reportId, "z ID:", reportFolder.getId());
    
    return reportFolder;
  } catch(e) {
    console.error("Błąd podczas tworzenia struktury folderów:", e);
    throw new Error("Nie można utworzyć folderów na zdjęcia: " + e.message);
  }
}

// Kompresja i przetwarzanie obrazu z formatu base64
function processImageData(base64Data) {
  try {
    // Usunięcie nagłówka z base64 (np. "data:image/jpeg;base64,")
    const base64Image = base64Data.split(',')[1];
    
    // Dekodowanie base64 do surowych danych binarnych
    const binaryData = Utilities.base64Decode(base64Image);
    
    // Konwersja do Blobu
    const blob = Utilities.newBlob(binaryData, MimeType.JPEG);
    
    return blob;
  } catch(e) {
    console.error("Błąd przetwarzania obrazu:", e);
    throw new Error("Nie można przetworzyć obrazu: " + e.message);
  }
}

// Zapisuje zdjęcia serwisowe dla danego raportu
function saveServiceImages(reportId, imagesData) {
  try {
    if (!imagesData || !imagesData.images || imagesData.images.length === 0) {
      console.log("Brak zdjęć do zapisania");
      return { success: false, message: "Brak zdjęć do zapisania" };
    }
      // Pobierz informacje o raporcie, aby uzyskać nazwę serwisu, numer ramy i nazwę operatora
    const reportInfo = getReportInfo(reportId);
    if (!reportInfo) {
      return { success: false, message: "Nie znaleziono informacji o raporcie: " + reportId };
    }
    
    // Utworzenie odpowiedniej struktury folderów
    const reportFolder = createReportImagesFolders(reportId);
    
    // Zapisywanie zdjęć
    const savedImages = [];
    const errors = [];
    
    // Aktualna data w formacie DD-MM-YYYY
    const currentDate = new Date();
    const formattedDate = `${String(currentDate.getDate()).padStart(2, '0')}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${currentDate.getFullYear()}`;
    
    imagesData.images.forEach((image, index) => {
      try {
        // Przetworzenie obrazu (kompresja)
        const processedImage = processImageData(image.data);        // Priorytetowo użyj pola frameNumber i serwisName przesłanego jako osobne pola
        let frameNumber = image.frameNumber || '';
        // Używamy nazwy serwisu zamiast operatora - najpierw sprawdzamy czy została przekazana z klienta
        let serviceName = image.serwisName || reportInfo.serwis || 'Serwis';
        
        // Jeśli frameNumber nie jest dostępne, spróbuj wyciągnąć z nazwy pliku
        if (!frameNumber && image.name) {
          const matches = image.name.match(/\[(SER\/[^[\]]+)\] - \[([^[\]]*)\] - \[([^[\]]*)\]/);
          if (matches && matches.length >= 3) {
            frameNumber = matches[2] || '';
            // Nie nadpisujemy operatora z nazwy pliku - użyjemy tego z reportInfo
          }
        }
        
        // Jeśli nadal nie mamy numeru ramy, użyj wartości z reportInfo
        if (!frameNumber) {
          frameNumber = reportInfo.bikeSerialNumber || '';
        }
          // Logowanie dla celów diagnostycznych
        console.log(`Przygotowywanie nazwy pliku - Numer ramy: "${frameNumber}", Serwis: "${serviceName}"`); 
        
        // Nowy format nazwy pliku: [ID raportu] - [Numer ramy] - [Nazwa serwisu] - DD-MM-YYYY - before/after X.jpg
        const filename = `[${reportId}] - [${frameNumber}] - [${serviceName}] - ${formattedDate} - ${image.type === 'before' ? 'before' : 'after'} ${index + 1}.jpg`;
        
        const file = reportFolder.createFile(processedImage.setName(filename));
        
        // Pobieranie URL do pliku
        const fileUrl = file.getUrl();
        console.log("Zapisano zdjęcie:", filename, "URL:", fileUrl);
        
        savedImages.push({
          name: filename,
          url: fileUrl,
          type: image.type
        });
      } catch (e) {
        console.error("Błąd zapisywania zdjęcia:", e);
        errors.push(`Błąd dla zdjęcia ${index + 1}: ${e.message}`);
      }
    });
    
    // Aktualizacja linków do zdjęć w raporcie
    if (savedImages.length > 0) {
      updateReportWithImageLinks(reportId, reportFolder.getUrl());
    }
    
    return {
      success: true,
      folderUrl: reportFolder.getUrl(),
      savedImages: savedImages,
      errors: errors.length > 0 ? errors : null
    };
  } catch(e) {
    console.error("Błąd zapisywania zdjęć serwisowych:", e);
    return { success: false, message: "Błąd zapisywania zdjęć: " + e.message };
  }
}

// Pobiera informacje o raporcie do uzupełnienia nazwy zdjęć
function getReportInfo(reportId) {
  try {
    console.log("Pobieranie informacji o raporcie:", reportId);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Sprawdzamy w nowym arkuszu "Raport serwisowy"
    let sheet = ss.getSheetByName('Raport serwisowy');
    
    // Jeśli nie znaleziono, sprawdź w starym arkuszu "Raporty serwisowe" dla kompatybilności wstecznej
    if (!sheet) {
      console.log("Nie znaleziono arkusza 'Raport serwisowy', sprawdzanie w 'Raporty serwisowe'");
      sheet = ss.getSheetByName('Raporty serwisowe');
    }
    
    if (!sheet) {
      console.error("Nie znaleziono żadnego arkusza z raportami");
      // Zwracamy wartość 'Serwis' zamiast 'Nieznany', co będzie bardziej opisowe w nazwie pliku
      return {
        serwis: 'Serwis',
        bikeSerialNumber: '',
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    console.log("Nagłówki w znalezionym arkuszu:", headers);
      // Indeksy kolumn - obsługuje różne nazwy kolumn
    const idColumnIndex = headers.indexOf('ID raportu');
    const serwisColumnIndex = headers.indexOf('Serwis');
    const operatorColumnIndex = headers.indexOf('Operator');
    
    // Sprawdzamy wszystkie możliwe nazwy kolumn dla numeru ramy
    let bikeSerialNumberColumnIndex = -1;
    const possibleFrameNumberColumns = ['Numer ramy', 'Numer seryjny', 'Bike Serial Number', 'bikeSerialNumber'];
    
    for (const columnName of possibleFrameNumberColumns) {
      const index = headers.indexOf(columnName);
      if (index !== -1) {
        bikeSerialNumberColumnIndex = index;
        console.log(`Znaleziono kolumnę z numerem ramy: ${columnName} (indeks: ${index})`);
        break;
      }
    }
      console.log("Indeksy kolumn:", {
      idColumnIndex: idColumnIndex,
      serwisColumnIndex: serwisColumnIndex,
      operatorColumnIndex: operatorColumnIndex,
      bikeSerialNumberColumnIndex: bikeSerialNumberColumnIndex
    });
    
    if (idColumnIndex === -1) {
      console.error("Nie znaleziono kolumny z ID raportu");
      return {
        serwis: 'Serwis',
        operator: '',
        bikeSerialNumber: '',
      };
    }
    
    // Używamy indeksów znalezionych kolumn lub domyślnych wartości
    const serwisIndex = serwisColumnIndex >= 0 ? serwisColumnIndex : 2; // Kolumna C
    const operatorIndex = operatorColumnIndex >= 0 ? operatorColumnIndex : 3; // Kolumna D w typowej strukturze
    const serialIndex = bikeSerialNumberColumnIndex >= 0 ? bikeSerialNumberColumnIndex : 6; // Kolumna G w nowej strukturze
    
    console.log(`Szukanie raportu o ID ${reportId} wśród ${data.length} wierszy...`);
    
    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][idColumnIndex];
        if (rowId === reportId) {
        const rowData = data[i];
        const reportInfo = {
          serwis: rowData[serwisIndex] || 'Serwis', 
          operator: rowData[operatorIndex] || '',
          bikeSerialNumber: rowData[serialIndex] || '',
        };
        
        console.log("Znaleziono informacje o raporcie:", reportInfo);
        return reportInfo;
      }
    }
    
    console.error("Nie znaleziono wiersza z ID raportu:", reportId);
    
    // Jeśli nie znaleziono raportu, zwróć podstawowe informacje
    return {
      serwis: 'Serwis',
      operator: '',
      bikeSerialNumber: '',
    };
  } catch(e) {
    console.error("Błąd pobierania danych raportu:", e);
    return {
      serwis: 'Serwis',
      bikeSerialNumber: '',
    };
  }
}

// Aktualizuje raport o linki do zdjęć
function updateReportWithImageLinks(reportId, folderUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Raport serwisowy');
    
    if (!sheet) {
      return false;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Znajdź indeks kolumny "Zdjęcia" według nagłówka
    let linkColumnIndex = headers.indexOf('Zdjęcia');
    if (linkColumnIndex === -1) {
      linkColumnIndex = 16; // fallback na Q
    }
    
    console.log("Aktualizacja raportu - używam kolumny Zdjęcia o indeksie:", linkColumnIndex);
    
    // Znajdź numer wiersza dla danego reportId
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reportId) {
        // Zaktualizuj link do zdjęć w kolumnie Zdjęcia
        sheet.getRange(i + 1, linkColumnIndex + 1).setValue(folderUrl);
        console.log("Zaktualizowano raport o link do zdjęć w kolumnie Zdjęcia:", folderUrl);
        return true;
      }
    }
    
    console.log("Nie znaleziono raportu do aktualizacji:", reportId);
    return false;
  } catch(e) {
    console.error("Błąd aktualizacji raportu o link do zdjęć:", e);
    return false;
  }
}

// Wysyła mail do klienta z potwierdzeniem raportu serwisowego
function sendReportEmailToClient(reportData) {
  try {
    // Hierarchia adresów email: 1) serviceConfirmationEmail z formularza, 2) Konfiguracja Email, 3) Użytkownicy, 4) Adresy
    let emailAddress = '';
    
    // POZIOM 1: Sprawdź email z formularza potwierdzenia serwisu (najwyższy priorytet)
    if (reportData.serviceConfirmationEmail) {
      emailAddress = reportData.serviceConfirmationEmail;
      console.log(`Używanie email z formularza potwierdzenia serwisu: ${emailAddress}`);
    }
    
    // POZIOM 2: Sprawdź konfigurację email (jeśli nie ma email z formularza)
    if (!emailAddress) {
      try {
        const emailConfig = getEmailConfiguration(reportData.serwis);
        if (emailConfig && emailConfig['Email Serwisu']) {
          emailAddress = emailConfig['Email Serwisu'];
          console.log(`Znaleziono adres email dla serwisu ${reportData.serwis} w konfiguracji email: ${emailAddress}`);
        }
      } catch (e) {
        console.error(`Błąd podczas szukania adresu email w konfiguracji email: ${e.message}`);
      }
    }
    
    // POZIOM 3: Jeśli nie znaleziono w konfiguracji, sprawdź zakładkę "Użytkownicy"
    if (!emailAddress) {
      try {
        const users = getDataFromSheet('Użytkownicy');
        // Szukamy użytkowników przypisanych do danego serwisu
        const serwisUsers = users.filter(u => u.Serwis === reportData.serwis);
        if (serwisUsers.length > 0 && serwisUsers[0].Email) {
          emailAddress = serwisUsers[0].Email;
          console.log(`Znaleziono adres email dla serwisu ${reportData.serwis} w zakładce Użytkownicy: ${emailAddress}`);
        }
      } catch (e) {
        console.error(`Błąd podczas szukania adresu email w zakładce Użytkownicy: ${e.message}`);
      }
    }
      // POZIOM 4: Jeśli nie znaleziono w zakładce Użytkownicy, spróbuj w zakładce Adresy
    if (!emailAddress) {
      try {
        const addresses = getDataFromSheet('Adresy');
        // Szukamy adresu dla serwisu z raportu
        const serwisAddress = addresses.find(a => a.Serwis === reportData.serwis);
        if (serwisAddress && serwisAddress.Email) {
          emailAddress = serwisAddress.Email;
          console.log(`Znaleziono adres email dla serwisu ${reportData.serwis} w zakładce Adresy: ${emailAddress}`);
        }
      } catch (e) {
        console.error(`Błąd podczas szukania adresu email w zakładce Adresy: ${e.message}`);
      }
    }
    
    // POZIOM 5: Jeśli nadal nie znaleziono, użyj adresu z danych raportu
    if (!emailAddress) {
      emailAddress = reportData.email || '';
      console.log(`Używanie adresu email z danych raportu: ${emailAddress}`);
    }
    
    // POZIOM 6: Jeśli nadal nie mamy adresu, spróbuj użyć danych z serviceAddressData
    if (!emailAddress) {
      const serviceAddressData = getServiceAddressData(reportData.serwis);
      if (serviceAddressData && serviceAddressData.email) {
        emailAddress = serviceAddressData.email;
        console.log(`Używanie adresu email z getServiceAddressData: ${emailAddress}`);
      }
    }
    
    if (!emailAddress) {
      console.error("Nie znaleziono adresu email do wysłania potwierdzenia");
      return false;
    }
    
    const subject = `Potwierdzenie raportu serwisowego nr ${reportData.id}`;
    
    // Przygotowanie tabeli z usługami - Z CENAMI dla serwisu
    let servicesTable = '';
    const serviceItems = reportData.serviceItems || [];
    
    if (serviceItems && serviceItems.length > 0) {
      servicesTable = `
        <div class="section">
          <div class="section-title">Usługi podczas naprawy</div>
          <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <thead>
              <tr style="background-color: #4285f4; color: white;">
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #3367d6;">Kod</th>
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #3367d6;">Nazwa usługi</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #3367d6;">Kategoria</th>
                <th style="padding: 12px; text-align: right; border-bottom: 2px solid #3367d6;">Cena</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      let totalServiceValue = 0;
      
      serviceItems.forEach((item, index) => {
        const rowBg = index % 2 === 0 ? '#f8f9fa' : '#ffffff';
        const customLabel = item.isCustom ? '<span style="display: inline-block; background-color: #fbbc04; color: #fff; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 5px;">niestandardowa</span>' : '';
        const itemPrice = parseFloat(item.price) || 0;
        totalServiceValue += itemPrice;
        
        servicesTable += `
          <tr style="background-color: ${rowBg};">
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #dddddd;">${item.code || ''}</td>
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #dddddd;">${item.name || ''} ${customLabel}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #dddddd;">${item.category || reportData.category || ''}</td>
            <td style="padding: 12px; text-align: right; border-bottom: 1px solid #dddddd;">${itemPrice.toFixed(2)} zł</td>
          </tr>
        `;
        
        // Dodaj opis dla niestandardowych usług, jeśli istnieje
        if (item.isCustom && item.description) {
          servicesTable += `
            <tr style="background-color: ${rowBg};">
              <td colspan="4" style="padding: 5px 12px 12px; text-align: left; border-bottom: 1px solid #dddddd; color: #666;">
                <em>Opis: ${item.description}</em>
              </td>
            </tr>
          `;
        }
      });
      
      // Dodaj wiersz z sumą
      servicesTable += `
        <tr style="background-color: #eaf3ff;">
          <td colspan="3" style="padding: 12px; text-align: right; border-bottom: 1px solid #dddddd; font-weight: bold;">Razem usługi:</td>
          <td style="padding: 12px; text-align: right; border-bottom: 1px solid #dddddd; font-weight: bold;">${totalServiceValue.toFixed(2)} zł</td>
        </tr>
      `;
      
      servicesTable += '</tbody></table></div>';
    }
    
    // Przygotowanie tabeli z częściami - BEZ CEN dla serwisu, tylko ilości
    let partsTable = '';
    const partItems = reportData.partItems || [];
    
    if (partItems && partItems.length > 0) {
      partsTable = `
        <div class="section">
          <div class="section-title">Części użyte do naprawy</div>
          <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <thead>
              <tr style="background-color: #4285f4; color: white;">
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #3367d6;">Kod</th>
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #3367d6;">Nazwa części</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #3367d6;">Kategoria</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #3367d6;">Ilość</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      partItems.forEach((item, index) => {
        const rowBg = index % 2 === 0 ? '#f8f9fa' : '#ffffff';
        const quantity = item.quantity || 1;
        
        partsTable += `
          <tr style="background-color: ${rowBg};">
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #dddddd;">${item.code || ''}</td>
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #dddddd;">${item.name || ''}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #dddddd;">${item.category || reportData.category || ''}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #dddddd;">${quantity}</td>
          </tr>
        `;
      });
      
      partsTable += '</tbody></table></div>';
    }
    
    // Dodanie sekcji z danymi roweru, tylko jeśli numer ramy jest określony
    let bikeDetailsSection = '';
    if (reportData.bikeSerialNumber && reportData.bikeSerialNumber !== '') {
      bikeDetailsSection = `
        <div class="section two-columns">
          <div class="column">
            <div class="section-title">Dane roweru</div>
            <p><span class="property-name">Numer seryjny:</span> <span class="property-value">${reportData.bikeSerialNumber}</span></p>
          </div>
        </div>
      `;
    }
    
    const body = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Google Sans', Roboto, Arial, sans-serif; line-height: 1.6; color: #202124; margin: 0; padding: 0; }
          .container { max-width: 650px; margin: 0 auto; }
          .header { background-color: #4285f4; color: white; padding: 30px 20px; text-align: center; }
          .content { background-color: #ffffff; padding: 30px; border-radius: 8px; margin-top: -20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .section { margin-bottom: 25px; padding-bottom: 20px; border-bottom: 1px solid #f1f3f4; }
          .section:last-child { border-bottom: none; }
          .section-title { font-size: 18px; font-weight: 500; margin-bottom: 15px; color: #4285f4; }
          .footer { margin-top: 30px; text-align: center; font-size: 13px; color: #5f6368; background-color: #f8f9fa; padding: 20px; border-radius: 0 0 8px 8px; }
          .order-id { font-size: 24px; font-weight: 500; margin: 15px 0; }
          .property-name { color: #5f6368; font-weight: 500; }
          .property-value { font-weight: normal; }
          .logo { font-size: 50px; margin-bottom: 15px; }
          .blue-text { color: #4285f4; }
          .top-banner { background-color: #1a73e8; height: 8px; }
          .two-columns { display: flex; flex-wrap: wrap; justify-content: space-between; }
          .column { width: 48%; }
          @media (max-width: 600px) {
            .column { width: 100%; margin-bottom: 20px; }
          }
        </style>
      </head>
      <body>
        <div class="top-banner"></div>
        <div class="container">
          <div class="header">
            <div class="logo">🔧</div>
            <h1>Potwierdzenie raportu serwisowego</h1>
            <div class="order-id">Raport nr: ${reportData.id}</div>
            <div>Data: ${reportData.serviceDate}</div>
          </div>
          <div class="content">
            <div class="section two-columns">
              <div class="column">
                <div class="section-title">Dane podstawowe</div>
                <p><span class="property-name">Serwis:</span> <span class="property-value">${reportData.serwis || ''}</span></p>
                <p><span class="property-name">Kategoria naprawy:</span> <span class="property-value">${reportData.category || ''}</span></p>
                <p><span class="property-name">Operator:</span> <span class="property-value">${reportData.operator || ''}</span></p>
                <p><span class="property-name">Oddział:</span> <span class="property-value">${reportData.department || ''}</span></p>
              </div>
              
              <div class="column">
                <div class="section-title">Szczegóły naprawy</div>
                <p><span class="property-name">Rodzaj serwisu:</span> <span class="property-value">${reportData.serviceType}</span></p>
                <p><span class="property-name">Data serwisu:</span> <span class="property-value">${reportData.serviceDate}</span></p>
                <p><span class="property-name">Przebieg (km):</span> <span class="property-value">${reportData.bikeMileage}</span></p>
                <p><span class="property-name">Roboczo godziny:</span> <span class="property-value">${reportData.laborHours}</span></p>
                <p><span class="property-name">Dojazd:</span> <span class="property-value">${reportData.isOnSite}</span></p>
              </div>
            </div>
            
            ${bikeDetailsSection}
            
            ${servicesTable}
            
            ${partsTable}
            
            ${reportData.notes ? `
            <div class="section">
              <div class="section-title">Uwagi do raportu</div>
              <p>${reportData.notes}</p>
            </div>
            ` : ''}
            
            <div class="footer">
              <p class="blue-text">Dziękujemy za skorzystanie z naszego serwisu.</p>
              <p>W przypadku pytań lub uwag prosimy o kontakt z serwisem.</p>
              <p>© ${new Date().getFullYear()} CARGO <span style="color: #2ecc71;">BIKE</span> SERWIS Sp. z o.o. | Kontakt: 📞 +48 888 788 847 | ✉️ biuro@cargobikeserwis.com</p>
            </div>
          </div>
        </div>
      </body>
      </html>
    `;
    
    // Wysyłanie emaila
    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: body
    });
    console.log(`Wysłano email z raportem do serwisu: ${emailAddress}`);
    return true;
  } catch(e) {
    console.error("Błąd wysyłania e-maila do klienta:", e);
    return false;
  }
}

// Wysyła mail do opiekuna serwisu z pełnymi danymi raportu (z cenami)
function sendMailToSupervisor(reportData, orderId) {
  try {
    // Hierarchia adresów email opiekuna: 1) Konfiguracja Email, 2) Arkusz Opiekun
    let supervisorEmail = '';
    let supervisorName = 'Opiekunie';
    
    // POZIOM 1: Sprawdź konfigurację email dla adresu opiekuna (najwyższy priorytet)
    try {
      const emailConfig = getEmailConfiguration(reportData.serwis);
      if (emailConfig && emailConfig['Email Opiekuna']) {
        supervisorEmail = emailConfig['Email Opiekuna'];
        console.log(`Znaleziono email opiekuna dla serwisu ${reportData.serwis} w konfiguracji email: ${supervisorEmail}`);
      }
    } catch (e) {
      console.error(`Błąd podczas szukania email opiekuna w konfiguracji email: ${e.message}`);
    }
    
    // POZIOM 2: Jeśli nie znaleziono w konfiguracji, użyj arkusza Opiekun
    let supervisor = null;
    if (!supervisorEmail) {
      supervisor = getDataFromSheet('Opiekun')[0];
      if (!supervisor || !supervisor.Email) {
        console.error("Brak danych opiekuna w arkuszu Opiekun");
        return false;
      }
      supervisorEmail = supervisor.Email;
      console.log(`Używanie email opiekuna z arkusza Opiekun: ${supervisorEmail}`);
    } else {
      // Jeśli mamy email z konfiguracji, pobierz dane opiekuna dla imienia
      try {
        supervisor = getDataFromSheet('Opiekun')[0];
      } catch (e) {
        console.log("Nie można pobrać danych opiekuna dla imienia, używanie domyślnego");
      }
    }
    
    // Pobierz imię opiekuna jeśli dane są dostępne
    if (supervisor) {
      // Debugowanie struktury opiekuna
      console.log('Pełna struktura obiektu opiekuna:', JSON.stringify(supervisor, null, 2));
      // Użyj dedykowanej funkcji do bezpiecznego pobierania imienia
      supervisorName = getSupervisorName(supervisor);
      console.log(`Końcowe imię opiekuna do użycia: ${supervisorName}`);
      
      // Dodatkowa walidacja - upewnij się, że nie jest undefined
      if (!supervisorName || supervisorName === 'undefined' || supervisorName.trim() === '') {
        console.error('UWAGA: Imię opiekuna jest nieprawidłowe, używanie domyślnego powitania');
        supervisorName = 'Opiekunie';
      }
    }
    
    if (!supervisorEmail) {
      console.error("Nie znaleziono adresu email opiekuna");
      return false;
    }
    
    const subject = `Raport serwisowy nr ${reportData.id || orderId} - ${reportData.serwis || 'Nieznany serwis'} - wymaga uwagi`;
    // Przygotowanie tabeli usług (z cenami)
    let servicesHtml = '';
    const serviceItems = reportData.services || [];
    let totalServiceValue = 0;
    if (serviceItems && serviceItems.length > 0) {
      servicesHtml = `
        <div class="section">
          <div class="section-title">Usługi podczas naprawy</div>
          <table style="width: 100%; border: 1px solid #b2dfdb; border-collapse: collapse; margin-top: 15px;">
            <thead>
              <tr style="background-color: #2ecc71; color: white;">
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #27ae60;">Kod</th>
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #27ae60;">Nazwa usługi</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #27ae60;">Kategoria</th>
                <th style="padding: 12px; text-align: right; border-bottom: 2px solid #27ae60;">Cena</th>
              </tr>
            </thead>
            <tbody>
      `;
      serviceItems.forEach((item, index) => {
        const rowBg = index % 2 === 0 ? '#eafaf1' : '#ffffff';
        const customLabel = item.isCustom ? '<span style="display: inline-block; background-color: #fbbc04; color: #fff; padding: 2px 6px; border-radius: 10px; font-size: 10px; margin-left: 5px;">niestandardowa</span>' : '';
        const itemPrice = parseFloat(item.price) || 0;
        totalServiceValue += itemPrice;
        servicesHtml += `
          <tr style="background-color: ${rowBg};">
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #b2dfdb;">${item.code || ''}</td>
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #b2dfdb;">${item.name || ''} ${customLabel}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #b2dfdb;">${item.category || reportData.category || ''}</td>
            <td style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb;">${itemPrice.toFixed(2)} zł</td>
          </tr>
        `;
        if (item.isCustom && item.description) {
          servicesHtml += `
            <tr style="background-color: ${rowBg};">
              <td colspan="4" style="padding: 5px 12px 12px; text-align: left; border-bottom: 1px solid #b2dfdb; color: #666;">
                <em>Opis: ${item.description}</em>
              </td>
            </tr>
          `;
        }
      });
      servicesHtml += `
        <tr style="background-color: #d4fbe8;">
          <td colspan="3" style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb; font-weight: bold;">Razem usługi:</td>
          <td style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb; font-weight: bold;">${totalServiceValue.toFixed(2)} zł</td>
        </tr>
      `;
      servicesHtml += '</tbody></table></div>';
    }
    // Przygotowanie tabeli części (z cenami)
    let partsHtml = '';
    const partItems = reportData.parts || [];
    let totalPartsValue = 0;
    if (partItems && partItems.length > 0) {
      partsHtml = `
        <div class="section">
          <div class="section-title">Części użyte do naprawy</div>
          <table style="width: 100%; border: 1px solid #b2dfdb; border-collapse: collapse; margin-top: 15px;">
            <thead>
              <tr style="background-color: #2ecc71; color: white;">
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #27ae60;">Kod</th>
                <th style="padding: 12px; text-align: left; border-bottom: 2px solid #27ae60;">Nazwa części</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #27ae60;">Kategoria</th>
                <th style="padding: 12px; text-align: center; border-bottom: 2px solid #27ae60;">Ilość</th>
                <th style="padding: 12px; text-align: right; border-bottom: 2px solid #27ae60;">Cena</th>
                <th style="padding: 12px; text-align: right; border-bottom: 2px solid #27ae60;">Wartość</th>
              </tr>
            </thead>
            <tbody>
      `;
      partItems.forEach((item, index) => {
        const rowBg = index % 2 === 0 ? '#eafaf1' : '#ffffff';
        const itemPrice = parseFloat(item.price) || 0;
        const quantity = parseInt(item.quantity) || 1;
        const itemValue = itemPrice * quantity;
        totalPartsValue += itemValue;
        partsHtml += `
          <tr style="background-color: ${rowBg};">
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #b2dfdb;">${item.code || ''}</td>
            <td style="padding: 12px; text-align: left; border-bottom: 1px solid #b2dfdb;">${item.name || ''}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #b2dfdb;">${item.category || reportData.category || ''}</td>
            <td style="padding: 12px; text-align: center; border-bottom: 1px solid #b2dfdb;">${quantity}</td>
            <td style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb;">${itemPrice.toFixed(2)} zł</td>
            <td style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb;">${itemValue.toFixed(2)} zł</td>
          </tr>
        `;
      });
      partsHtml += `
        <tr style="background-color: #d4fbe8;">
          <td colspan="5" style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb; font-weight: bold;">Razem części:</td>
          <td style="padding: 12px; text-align: right; border-bottom: 1px solid #b2dfdb; font-weight: bold;">${totalPartsValue.toFixed(2)} zł</td>
        </tr>
      `;
      partsHtml += '</tbody></table></div>';
    }
    // Sekcja podsumowania finansowego (zawsze widoczna)
    const summaryHtml = `
      <div class="section" style="background:#eafaf1; border-left:4px solid #2ecc71; margin-top:20px; padding:16px;">
        <div class="section-title" style="color:#27ae60;">Podsumowanie finansowe</div>
        <table style="width:100%; font-size:16px;">
          <tr>
            <td style="padding:8px 0;">Suma usług:</td>
            <td style="padding:8px 0; text-align:right; font-weight:bold;">${totalServiceValue.toFixed(2)} zł</td>
          </tr>
          <tr>
            <td style="padding:8px 0;">Suma części:</td>
            <td style="padding:8px 0; text-align:right; font-weight:bold;">${totalPartsValue.toFixed(2)} zł</td>
          </tr>
          <tr>
            <td style="padding:8px 0; font-size:18px;">Suma całkowita:</td>
            <td style="padding:8px 0; text-align:right; font-size:18px; color:#27ae60; font-weight:bold;">${(totalServiceValue+totalPartsValue).toFixed(2)} zł</td>
          </tr>
        </table>
      </div>
    `;
    const body = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Google Sans', Roboto, Arial, sans-serif; line-height: 1.6; color: #202124; margin: 0; padding: 0; }
          .container { max-width: 800px; margin: 0 auto; }
          .header { background-color: #2ecc71; color: white; padding: 30px 20px; text-align: center; }
          .content { background-color: #ffffff; padding: 30px; border-radius: 8px; margin-top: -20px; box-shadow: 0 2px 10px rgba(46,204,113,0.1); }
          .section { margin-bottom: 25px; padding-bottom: 20px; border-bottom: 1px solid #eafaf1; }
          .section:last-child { border-bottom: none; }
          .section-title { font-size: 18px; font-weight: 500; margin-bottom: 15px; color: #27ae60; }
          .footer { margin-top: 30px; text-align: center; font-size: 13px; color: #5f6368; background-color: #f8f9fa; padding: 20px; border-radius: 0 0 8px 8px; }
          .report-id { font-size: 24px; font-weight: 500; margin: 15px 0; }
          .property-name { color: #5f6368; font-weight: 500; }
          .property-value { font-weight: normal; }
          .logo { font-size: 50px; margin-bottom: 15px; }
          .blue-text { color: #4285f4; }
          .green-text { color: #0f9d58; }
          .top-banner { background-color: #2ecc71; height: 8px; }
          .two-columns { display: flex; flex-wrap: wrap; justify-content: space-between; }
          .column { width: 48%; }
          @media (max-width: 600px) {
            .column { width: 100%; margin-bottom: 20px; }
          }
        </style>
      </head>
      <body>
        <div class="top-banner"></div>
        <div class="container">
          <div class="header">            <div class="logo">🔧</div>
            <h1>Potwierdzenie raportu serwisowego</h1>
            <div class="report-id">Raport nr: ${reportData.id}</div>
            <div>Data: ${reportData.serviceDate}</div>
            <div>Witaj ${supervisorName},</div>
          </div>
          <div class="content">
            <div class="section two-columns">
              <div class="column">
                <div class="section-title">Dane podstawowe</div>
                <p><span class="property-name">Serwis:</span> <span class="property-value">${reportData.serwis || ''}</span></p>
                <p><span class="property-name">Kategoria naprawy:</span> <span class="property-value">${reportData.category || ''}</span></p>
                <p><span class="property-name">Operator:</span> <span class="property-value">${reportData.operator || ''}</span></p>
                <p><span class="property-name">Oddział:</span> <span class="property-value">${reportData.department || ''}</span></p>
              </div>
              <div class="column">
                <div class="section-title">Szczegóły naprawy</div>
                <p><span class="property-name">Rodzaj serwisu:</span> <span class="property-value">${reportData.serviceType}</span></p>
                <p><span class="property-name">Data serwisu:</span> <span class="property-value">${reportData.serviceDate}</span></p>
                <p><span class="property-name">Przebieg (km):</span> <span class="property-value">${reportData.bikeMileage}</span></p>
                <p><span class="property-name">Roboczo godziny:</span> <span class="property-value">${reportData.laborHours}</span></p>
                <p><span class="property-name">Dojazd:</span> <span class="property-value">${reportData.isOnSite}</span></p>
              </div>
            </div>
            
            ${servicesHtml}
            ${partsHtml}
            ${summaryHtml}
            ${reportData.notes ? `
            <div class="section">
              <div class="section-title">Uwagi do raportu</div>
              <p>${reportData.notes}</p>
            </div>
            ` : ''}
            <div class="footer">
              <p class="green-text">Dziękujemy za skorzystanie z naszego serwisu.</p>
              <p>W przypadku pytań lub uwag prosimy o kontakt z serwisem.</p>
              <p>© ${new Date().getFullYear()} CARGO <span style="color: #2ecc71;">BIKE</span> SERWIS Sp. z o.o. | Kontakt: 📞 +48 888 788 847 | ✉️ biuro@cargobikeserwis.com</p>
            </div>
          </div>
        </div>
      </body>
      </html>    `;
    MailApp.sendEmail({
      to: supervisorEmail,
      subject: subject,
      htmlBody: body
    });
    console.log(`Wysłano email z raportem do opiekuna: ${supervisorEmail}`);
    return true;
  } catch(e) {
    console.error("Błąd wysyłania e-maila do opiekuna:", e);
    return false;
  }
}

// Funkcja testowa do weryfikacji poprawności pobierania imienia opiekuna
function testSupervisorNameRetrieval() {
  console.log('=== TEST POBIERANIA IMIENIA OPIEKUNA ===');
  
  // Test z różnymi strukturami danych opiekuna
  const testCases = [
    // Case 1: Standard 'Imię' field
    { Imię: 'Jan Kowalski', Email: 'opiekun@example.com' },
    // Case 2: Alternative 'Imie' field (without accent)
    { Imie: 'Anna Nowak', Email: 'opiekun2@example.com' },
    // Case 3: English 'Name' field
    { Name: 'John Smith', Email: 'supervisor@example.com' },
    // Case 4: Only first column without header (edge case)
    { '': 'Test User', Email: 'test@example.com' },
    // Case 5: Empty supervisor object
    {},
    // Case 6: No name field at all
    { Email: 'only@email.com', SomeOtherField: 'value' }
  ];
  
  testCases.forEach((testCase, index) => {
    console.log(`\n--- Test Case ${index + 1} ---`);
    console.log('Dane wejściowe:', JSON.stringify(testCase, null, 2));
    
    const result = getSupervisorName(testCase);
    console.log(`Wynik: "${result}"`);
    
    // Validate result
    if (result === 'undefined' || result === undefined) {
      console.error('BŁĄD: Funkcja zwróciła undefined!');
    } else {
      console.log('✓ Funkcja działała poprawnie');
    }
  });
  
  console.log('\n=== KONIEC TESTÓW ===');
  return 'Testy zakończone - sprawdź logi w konsoli';
}

// Bezpieczne pobieranie imienia opiekuna z różnymi wariantami kolumn
function getSupervisorName(supervisor) {
  if (!supervisor) {
    console.log('Brak obiektu opiekuna');
    return 'Opiekunie';
  }
  
  // Lista możliwych nazw kolumn dla imienia opiekuna
  const possibleNameFields = [
    'Imię', 'Imie', 'Name', 'name', 'Nazwa', 'nazwa',
    'FirstName', 'firstName', 'first_name', 'IMIĘ', 'IMIE',
    'Pełne imię', 'Pełne Imię', 'Imię i nazwisko', 'Imię i Nazwisko'
  ];
  
  console.log('Dostępne klucze w obiekcie opiekuna:', Object.keys(supervisor));
  
  // Sprawdź każde możliwe pole
  for (const field of possibleNameFields) {
    const value = supervisor[field];
    if (value && typeof value === 'string' && value.trim() !== '' && value.toLowerCase() !== 'undefined') {
      console.log(`Znaleziono imię opiekuna w polu '${field}': ${value}`);
      return value.trim();
    }
  }
  
  // Jeśli nie znaleziono żadnego imienia, sprawdź pierwszą kolumnę (może być bez nagłówka)
  const keys = Object.keys(supervisor);
  if (keys.length > 0) {
    const firstValue = supervisor[keys[0]];
    if (firstValue && typeof firstValue === 'string' && firstValue.trim() !== '' && 
        firstValue !== supervisor.Email && firstValue.toLowerCase() !== 'undefined') {
      console.log(`Używanie pierwszej kolumny '${keys[0]}' jako imienia opiekuna: ${firstValue}`);
      return firstValue.trim();
    }
  }
  
  console.log('Nie znaleziono prawidłowego imienia opiekuna, używanie domyślnego powitania');
  return 'Opiekunie';
}

// Funkcja do zarządzania konfiguracją adresów email dla serwisów
function getEmailConfiguration(serwis) {
  try {
    const emailConfigs = getDataFromSheet('Konfiguracja Email');
    return emailConfigs.find(config => config.Serwis === serwis);
  } catch (e) {
    console.log("Arkusz 'Konfiguracja Email' nie istnieje");
    return null;
  }
}

// Funkcja do zapisywania konfiguracji email dla serwisu
function saveEmailConfiguration(serwis, emailSerwisu, emailOpiekuna) {
  try {
    let emailConfigs = [];
    let sheet;
    
    try {
      emailConfigs = getDataFromSheet('Konfiguracja Email');
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Konfiguracja Email');
    } catch (e) {
      // Arkusz nie istnieje, utwórz go
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Konfiguracja Email');
      sheet.getRange(1, 1, 1, 3).setValues([['Serwis', 'Email Serwisu', 'Email Opiekuna']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      emailConfigs = [];
    }
    
    // Sprawdź czy konfiguracja dla tego serwisu już istnieje
    const existingIndex = emailConfigs.findIndex(config => config.Serwis === serwis);
    
    if (existingIndex !== -1) {
      // Aktualizuj istniejącą konfigurację
      const rowIndex = existingIndex + 2; // +2 bo nagłówek w rzędzie 1, a indeksy od 0
      sheet.getRange(rowIndex, 2).setValue(emailSerwisu || '');
      sheet.getRange(rowIndex, 3).setValue(emailOpiekuna || '');
    } else {
      // Dodaj nową konfigurację
      const newRow = emailConfigs.length + 2; // +2 bo nagłówek w rzędzie 1
      sheet.getRange(newRow, 1, 1, 3).setValues([[serwis, emailSerwisu || '', emailOpiekuna || '']]);
    }
    
    return { success: true, message: "Konfiguracja email została zapisana" };
  } catch (e) {
    console.error("Błąd zapisywania konfiguracji email:", e);
    return { success: false, message: "Błąd zapisywania konfiguracji: " + e.message };
  }
}

// Funkcja do pobierania wszystkich konfiguracji emaili
function getAllEmailConfigurations() {
  try {
    const emailConfigs = getDataFromSheet('Konfiguracja Email');
    return { success: true, configurations: emailConfigs };
  } catch (e) {
    console.log("Arkusz 'Konfiguracja Email' nie istnieje");
    return { success: true, configurations: [] };
  }
}
