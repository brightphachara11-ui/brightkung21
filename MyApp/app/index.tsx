import { useEffect, useState } from 'react';
import { View, Text, StyleSheet, Dimensions, TouchableOpacity, Alert, ScrollView } from 'react-native';
import { Image } from 'expo-image';
import * as MediaLibrary from 'expo-media-library';
import { GestureHandlerRootView, GestureDetector, Gesture } from 'react-native-gesture-handler';
import Animated, { useSharedValue, useAnimatedStyle, withSpring, runOnJS } from 'react-native-reanimated';

const { width, height } = Dimensions.get('window');

export default function App() {
  const [photos, setPhotos] = useState([]);
  const [current, setCurrent] = useState(0);
  const [toDelete, setToDelete] = useState([]);
  const [status, setStatus] = useState('Loading...');
  const [reviewing, setReviewing] = useState(false);
  const [permission, requestPermission] = MediaLibrary.usePermissions();
  const translateX = useSharedValue(0);
  const rotate = useSharedValue(0);

  useEffect(() => {
    if (permission?.granted) loadPhotos();
    else requestPermission();
  }, [permission]);

  const loadPhotos = async () => {
    try {
      const media = await MediaLibrary.getAssetsAsync({ mediaType: 'photo', first: 100 });
      const shuffled = media.assets.sort(() => Math.random() - 0.5);
      setPhotos(shuffled);
      setStatus('');
    } catch (e) {
      setStatus('Error: ' + e.message);
    }
  };

  const formatSize = (bytes) => {
    if (!bytes) return '? KB';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const totalSize = (list) => {
    const total = list.reduce((sum, p) => sum + (p.fileSize || 0), 0);
    if (total < 1024 * 1024) return (total / 1024).toFixed(1) + ' KB';
    return (total / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const goNext = () => {
    translateX.value = withSpring(0);
    rotate.value = withSpring(0);
    setCurrent(c => {
      if (c + 1 >= photos.length) setReviewing(true);
      return c + 1;
    });
  };

  const markDelete = () => {
    const photo = photos[current];
    setToDelete(prev => [...prev, photo]);
    translateX.value = withSpring(-width * 1.5);
    setCurrent(c => {
      if (c + 1 >= photos.length) setReviewing(true);
      return c + 1;
    });
  };

  const gesture = Gesture.Pan()
    .onUpdate((e) => {
      translateX.value = e.translationX;
      rotate.value = e.translationX / 20;
    })
    .onEnd((e) => {
      if (e.translationX < -100) {
        runOnJS(markDelete)();
      } else if (e.translationX > 100) {
        runOnJS(goNext)();
      } else {
        translateX.value = withSpring(0);
        rotate.value = withSpring(0);
      }
    });

  const animatedStyle = useAnimatedStyle(() => ({
    transform: [
      { translateX: translateX.value },
      { rotate: `${rotate.value}deg` }
    ],
  }));

  const deleteOpacity = useAnimatedStyle(() => ({
    opacity: translateX.value < -50 ? Math.min(1, Math.abs(translateX.value) / 100) : 0
  }));

  const keepOpacity = useAnimatedStyle(() => ({
    opacity: translateX.value > 50 ? Math.min(1, translateX.value / 100) : 0
  }));

  const confirmDeleteAll = async () => {
    Alert.alert(
      'Delete All?',
      `Permanently delete ${toDelete.length} photos (${totalSize(toDelete)})?`,
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Delete All', style: 'destructive', onPress: async () => {
            try {
              await MediaLibrary.deleteAssetsAsync(toDelete.map(p => p.id));
              Alert.alert('Done!', `${toDelete.length} photos deleted!`);
              setToDelete([]);
              setReviewing(false);
              setCurrent(0);
              loadPhotos();
            } catch (e) {
              Alert.alert('Error', e.message);
            }
          }
        }
      ]
    );
  };

  const removeFromDeleteList = (id) => {
    setToDelete(prev => prev.filter(p => p.id !== id));
  };

  const photo = photos[current];

  // Review screen
  if (reviewing || current >= photos.length) {
    return (
      <View style={styles.container}>
        <Text style={styles.counter}>Review Delete List</Text>
        <Text style={styles.size}>
          {toDelete.length} photos • {totalSize(toDelete)} will be freed
        </Text>

        {toDelete.length === 0 ? (
          <Text style={styles.status}>No photos marked for deletion!</Text>
        ) : (
          <ScrollView style={styles.grid} contentContainerStyle={styles.gridContent}>
            {toDelete.map((p) => (
              <View key={p.id} style={styles.gridItem}>
                <Image source={p.uri} style={styles.thumb} contentFit="cover" />
                <Text style={styles.thumbSize}>{formatSize(p.fileSize)}</Text>
                <TouchableOpacity style={styles.removeBtn} onPress={() => removeFromDeleteList(p.id)}>
                  <Text style={styles.removeBtnText}>↩ Keep</Text>
                </TouchableOpacity>
              </View>
            ))}
          </ScrollView>
        )}

        <View style={styles.buttons}>
          {toDelete.length > 0 && (
            <TouchableOpacity style={styles.deleteBtn} onPress={confirmDeleteAll}>
              <Text style={styles.btnText}>🗑 Delete All ({toDelete.length})</Text>
            </TouchableOpacity>
          )}
          <TouchableOpacity style={styles.keepBtn} onPress={() => { setReviewing(false); setCurrent(0); loadPhotos(); }}>
            <Text style={styles.btnText}>🔄 Start Over</Text>
          </TouchableOpacity>
        </View>
      </View>
    );
  }

  return (
    <GestureHandlerRootView style={styles.container}>
      {status ? (
        <Text style={styles.status}>{status}</Text>
      ) : photo ? (
        <>
          <Text style={styles.counter}>{current + 1} / {photos.length} • 🗑 {toDelete.length}</Text>

          <Animated.View style={[styles.label, styles.deleteLabel, deleteOpacity]}>
            <Text style={styles.labelText}>🗑 DELETE</Text>
          </Animated.View>

          <Animated.View style={[styles.label, styles.keepLabel, keepOpacity]}>
            <Text style={styles.labelText}>✅ KEEP</Text>
          </Animated.View>

          <GestureDetector gesture={gesture}>
            <Animated.View style={animatedStyle}>
              <Image source={photo.uri} style={styles.image} contentFit="cover" />
            </Animated.View>
          </GestureDetector>

          <Text style={styles.size}>📁 {formatSize(photo.fileSize)}</Text>

          <View style={styles.buttons}>
            <TouchableOpacity style={styles.deleteBtn} onPress={markDelete}>
              <Text style={styles.btnText}>🗑 Delete</Text>
            </TouchableOpacity>
            <TouchableOpacity style={styles.keepBtn} onPress={goNext}>
              <Text style={styles.btnText}>✅ Keep</Text>
            </TouchableOpacity>
          </View>
        </>
      ) : null}
    </GestureHandlerRootView>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#000', alignItems: 'center', justifyContent: 'center' },
  counter: { color: '#aaa', fontSize: 14, marginBottom: 10 },
  image: { width: width, height: height * 0.7 },
  size: { color: '#fff', fontSize: 16, marginTop: 10 },
  status: { color: '#fff', fontSize: 18, textAlign: 'center', padding: 20 },
  buttons: { flexDirection: 'row', gap: 20, marginTop: 15, flexWrap: 'wrap', justifyContent: 'center' },
  deleteBtn: { backgroundColor: '#ff4444', padding: 15, borderRadius: 10 },
  keepBtn: { backgroundColor: '#44bb44', padding: 15, borderRadius: 10 },
  btnText: { color: '#fff', fontSize: 16, fontWeight: 'bold' },
  label: { position: 'absolute', top: 100, zIndex: 10, padding: 10, borderRadius: 10, borderWidth: 3 },
  deleteLabel: { left: 30, borderColor: '#ff4444' },
  keepLabel: { right: 30, borderColor: '#44bb44' },
  labelText: { fontSize: 24, fontWeight: 'bold', color: '#fff' },
  grid: { width: width, maxHeight: height * 0.6 },
  gridContent: { flexDirection: 'row', flexWrap: 'wrap', padding: 5 },
  gridItem: { width: width / 3 - 10, margin: 5, alignItems: 'center' },
  thumb: { width: width / 3 - 10, height: width / 3 - 10, borderRadius: 8 },
  thumbSize: { color: '#aaa', fontSize: 10, marginTop: 3 },
  removeBtn: { backgroundColor: '#444', padding: 4, borderRadius: 5, marginTop: 3 },
  removeBtnText: { color: '#fff', fontSize: 11 },
});